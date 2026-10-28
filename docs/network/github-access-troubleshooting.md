# 一次 GitHub 访问排查：从 22 秒超时到 SSH over 443

> 记录一次真实的网络排障过程。全部数据来自本机实测，方法可复用到任何"连接得上但用不了"的问题。

---

## 一、现象

访问 `https://github.com` 极慢，页面偶尔能打开，但 `git push` 基本不可用。

用 `curl` 的分段时间指标一测，问题就很具体了：

```powershell
curl.exe -o NUL -s -S -w "DNS=%{time_namelookup}s`nTCP=%{time_connect}s`nTLS=%{time_appconnect}s`n首字节=%{time_starttransfer}s`n总计=%{time_total}s`n" https://github.com/
```

```
DNS=0.044799s
TCP=0.127552s
TLS=0.888962s
首字节=0.000000s
总计=22.559300s
```

---

## 二、分层判读

`curl` 的这几个指标对应网络栈的不同层次，把总耗时拆开，就能看出卡在哪一层：

| 阶段 | 指标 | 实测 | 判断 |
|---|---|---|---|
| DNS 解析 | `time_namelookup` | 0.045s | ✅ 正常，不是 DNS 问题 |
| TCP 握手 | `time_connect` | 0.128s | ✅ 正常，端口通、能建连 |
| TLS 握手 | `time_appconnect` | 0.889s | ⚠️ 偏慢，但握手完成了 |
| **首字节** | `time_starttransfer` | **0.000s** | ❌ **一个字节都没收到** |
| 总计 | `time_total` | 22.559s | ❌ 卡了 21 秒后放弃 |

### 三个关键推论

**1. 不是 DNS 污染。**
DNS 只花了 45ms，解析是成功的。所以改 DNS、写 hosts 这类方案对本问题无效——它们解决的是"解析到哪个 IP"，而问题出在"到了那个 IP 之后拿不到数据"。

**2. 不是服务器慢。**
如果是服务器处理慢，首字节会是 5s、10s 这种"有响应但很迟"的值。
**首字节为 0 意味着请求发出去了，但一个响应字节都没回来。** 这排除了"服务端性能问题"。

**3. 不是端口被封。**
TCP 握手 128ms 完成，TLS 握手也完成了（0.889s）。说明 443 端口能连，证书协商也走完了。**被卡住的是握手之后的数据传输阶段。**

**4. 22.5 秒这个数字本身是线索。**
21 秒多的等待，符合 **TCP 数据重传的指数退避**（约 1+2+4+8+… 的累积）。说明数据包发出后**被静默丢弃**，TCP 一直在重传，直到放弃——而不是收到 RST 快速失败。

---

## 三、排除过程

按"先便宜、后昂贵"的顺序逐层排除：

| # | 怀疑对象 | 怎么测 | 结果 | 结论 |
|---|---|---|---|---|
| 1 | DNS 解析 | `nslookup github.com` vs `nslookup github.com 223.5.5.5` | 一致，45ms | 排除 |
| 2 | 端口被封 | `Test-NetConnection github.com -Port 443` | 成功 | 排除 |
| 3 | 本地线路/运营商 | 手机热点重测同一条命令 | 一样慢 | **排除** |
| 4 | 公司网络策略 | 在家庭网络复测 | 一样慢 | 排除 |
| 5 | 服务器性能 | `time_starttransfer = 0` | 未收到响应 | 排除 |
| 6 | HTTP 路径本身 | 换一个入口测试 | 见下 | **定位** |

第 3 步是关键：**换一条完全不同的网络（手机热点），现象不变**，说明问题不在自己的宽带和运营商。

到这一步，"连接建立成功、但数据回不来"已经能确定，剩下的问题变成：**换个路径能不能通。**

---

## 四、验证：换一个入口

GitHub 除了 `github.com`（140.82.x 段），还有 `ssh.github.com`（20.205.243.x 段）——**不同的一组边缘节点**。

```powershell
Test-NetConnection ssh.github.com -Port 443
```

```
ComputerName     : ssh.github.com
RemoteAddress    : 20.205.243.160
RemotePort       : 443
TcpTestSucceeded : True
```

TCP 通。但 TCP 通不等于数据能走，继续测协议层：

```powershell
ssh -T -p 443 git@ssh.github.com
```

```
The authenticity of host '[ssh.github.com]:443 ([20.205.243.160]:443)' can't be established.
ED25519 key fingerprint is SHA256:+DiY3wvvV6TuJJhbpZisF/zLDA0zPMSvHdkr4UvCOqU.
Are you sure you want to continue connecting (yes/no/[fingerprint])? yes
Warning: Permanently added '[ssh.github.com]:443' (ED25519) to the list of known hosts.
git@ssh.github.com: Permission denied (publickey).
```

**`Permission denied (publickey)` 在这里是好消息。**

它说明：TCP 通了 → SSH 协议协商成功 → 服务端正常返回响应 → **整条链路双向都能走数据**，只是还没配公钥。

对比 `github.com` 那条路（首字节 0、卡 21 秒），差距一目了然：**同一台机器、同一个出口，换个入口就通了。**

### 顺手校验主机指纹

`yes` 之前应该核对指纹。GitHub 官方公布的是（[文档](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/githubs-ssh-key-fingerprints)）：

| 类型 | 官方指纹 |
|---|---|
| Ed25519 | `SHA256:+DiY3wvvV6TuJJhbpZisF/zLDA0zPMSvHdkr4UvCOqU` |
| ECDSA | `SHA256:p2QAMXNIC1TJYWeIOttrVc98/R1BUFWu3/LiyKgUfQM` |
| RSA | `SHA256:uNiVztksCsDhcc0u9e8BujQXVUpKZIDTMczCvj3tD2s` |

实测值与官方 Ed25519 指纹一致，连接可信。

> 这一步不能省。第一次连陌生主机时输 `yes` 之前先对指纹，和浏览器看 HTTPS 证书是同一件事——防的是中间人。

---

## 五、最终方案

### 1. 生成密钥

```powershell
ssh-keygen -t ed25519 -C "你的邮箱"
# 三次回车：默认路径、密码留空、再确认
```

### 2. 把公钥加到 GitHub

```powershell
Get-Content ~\.ssh\id_ed25519.pub | Set-Clipboard
# 打开 https://github.com/settings/keys → New SSH key → 粘贴
```

### 3. 让 git 默认走 443

`~/.ssh/config`：

```
Host github.com
    HostName ssh.github.com
    Port 443
    User git
    IdentityFile ~/.ssh/id_ed25519
    ServerAliveInterval 30
    ServerAliveCountMax 3
```

这样写 `git@github.com:user/repo.git` 时会自动转成 `ssh.github.com:443`，不用每次加 `-p 443`。

### 4. 验证

```powershell
ssh -T git@github.com
```

```
Hi <用户名>! You've successfully authenticated, but GitHub does not provide shell access.
```

---

## 六、过程中的两个真实报错

### `GH007: Your push would publish a private email address`

提交作者用的是真实邮箱，而 GitHub 开了"阻止命令行推送暴露邮箱"的保护。

```powershell
git config user.email "<数字ID>+<用户名>@users.noreply.github.com"
git commit --amend --reset-author --no-edit
git push -u origin main
```

> 教训：**身份要在第一次 commit 之前配好。** 只有一个提交时改起来是一条命令，有二十个提交就得重写历史。

### `src refspec main does not match any`

`git init` 之后仓库是空的，`HEAD` 指向 `refs/heads/master` 但这个 ref 还不存在（要等第一次 commit 才创建）。所以先 `git branch -M main` 会失败。

正确顺序：**先 commit，再改名。**

---

## 七、可复用的方法

这套排查的价值不在"GitHub 怎么访问"，而在下面这个通用套路：

```
1. 量化现象     用分段时间指标把"慢"拆成 DNS / TCP / TLS / 首字节
2. 逐层排除     每一层给出"正常/异常"的判据，而不是靠感觉
3. 更换变量     换网络、换入口、换协议，每次只改一个
4. 验证假设     能通的那个组合，就是定位结果
5. 固化方案     把可用的路径写成配置，不再依赖记忆
```

**判据比结论重要。** 比如"首字节为 0 而不是 5s"这个区别，直接分开了"服务端慢"和"数据回不来"两种完全不同的故障——如果只看总耗时 22 秒，是分不出来的。

---

## 八、遗留问题

- **没有做 SNI 对照实验**（同一个 IP 换一个域名做 TLS）。因此"是按域名干扰"只是推断，不是结论。若要确认，可用：

  ```powershell
  $ip = (Resolve-DnsName github.com -Type A | Where-Object {$_.IPAddress}).IPAddress[0]
  curl.exe -k -s -o NUL -m 12 -w "同IP换域名 总=%{time_total}s`n" --resolve "example.com:443:$ip" https://example.com/
  curl.exe -k -s -o NUL -m 12 -w "同IP同域名 总=%{time_total}s`n" --resolve "github.com:443:$ip" https://github.com/
  ```

  若前者快、后者慢，说明干扰按域名生效，那么换 IP / 改 hosts 全部无效，只能走 SSH 或代理。

- **网页访问仍然慢。** 但 git 的推拉走 SSH 443，已经和网页解耦，日常不受影响。

---

> 本文所有数据均为本机真实测量。主机名、IP 均为 GitHub 公开地址，不含任何内部信息。

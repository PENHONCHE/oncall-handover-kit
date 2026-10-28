@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   值班视图演示（无需 Docker）
echo ============================================
echo.
echo 正在启动假指标生成器（端口 9100）...

start "oncall-demo-exporter" python mock_exporter.py --port 9100
timeout /t 3 >nul

echo 正在打开看板预览...
start "" "preview.html"

echo.
echo 底部四个按钮按顺序点：正常 -^> 延迟劣化 -^> 链路中断 -^> 恢复
echo 这就是一次真实网络故障的完整形状。
echo.
echo 关闭演示：直接关掉标题为 oncall-demo-exporter 的窗口。
echo.
pause

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
值班视图演示用指标生成器（纯标准库，零第三方依赖）

目的：在完全脱离公司环境的前提下，复现"值班大屏看不到故障域"这个问题，
      并演示"如果看板上多了链路状态和依赖关系，值班的人会看到什么"。

运行：
    python mock_exporter.py --port 9100

打开 http://localhost:9100/state 可查看当前场景。

控制故障场景（用浏览器或 curl 访问）：
    /control?scenario=normal    正常（默认）
    /control?scenario=degrade   延迟劣化：P90/P99 上升、误码增长、光功率下降，但链路仍 up
    /control?scenario=down      链路中断：连接数归零、主备链路同时 down
    /control?scenario=recover   恢复

演示剧本（复现一次真实故障的形状）：
    degrade  →  down  →  recover

注意：本文件生成的每一个数字都是随机模拟的，不含任何真实数据。
"""

import argparse
import json
import math
import random
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

START = time.time()

STATE = {
    "scenario": "normal",
    "changed_at": time.time(),
}
LOCK = threading.Lock()

# 计数器需要单调递增，单独维护
COUNTERS = {
    "crc_primary": 12_000.0,
    "crc_backup": 11_800.0,
}

SCENARIOS = ("normal", "degrade", "down", "recover")


def _noise(base, ratio=0.03):
    """给指标加一点抖动，避免曲线是一条直线（看起来才像真的）。"""
    return base * (1.0 + random.uniform(-ratio, ratio))


def _since_change():
    with LOCK:
        return time.time() - STATE["changed_at"]


def _current_scenario():
    with LOCK:
        return STATE["scenario"]


def collect():
    """根据当前场景生成一组指标快照。"""
    scenario = _current_scenario()
    elapsed = _since_change()

    # 默认：正常
    availability = 1.0
    connections = _noise(1240, 0.06)
    p75, p90, p99 = _noise(0.082), _noise(0.151), _noise(0.342)
    rps = _noise(860, 0.08)
    link_primary = 1.0
    link_backup = 1.0
    rx_primary = _noise(-12.4, 0.02)
    rx_backup = _noise(-12.9, 0.02)
    crc_rate = 0.4          # 每秒误码增长
    incident_open = 0.0
    pending_items = 3.0

    if scenario == "degrade":
        # 物理层开始劣化：链路还在 up，但误码飙升、光功率下降
        progress = min(elapsed / 20.0, 1.0)      # 20 秒内逐步恶化，方便演示时观察
        availability = 1.0
        connections = _noise(1240 * (1 - 0.25 * progress), 0.08)
        p75 = _noise(0.082 + 0.35 * progress)
        p90 = _noise(0.151 + 1.05 * progress)
        p99 = _noise(0.342 + 3.10 * progress)
        rps = _noise(860 * (1 - 0.3 * progress), 0.10)
        link_primary = 1.0
        link_backup = 1.0
        rx_primary = _noise(-12.4 - 9.0 * progress, 0.03)
        rx_backup = _noise(-12.9 - 8.6 * progress, 0.03)
        crc_rate = 900.0 * progress
        incident_open = 1.0
        pending_items = 5.0

    elif scenario == "down":
        # 链路彻底中断：主备同时 down，业务不可用
        availability = 0.0
        connections = 0.0
        p75 = p90 = p99 = 0.0        # 没有完成请求，分位数没有意义
        rps = 0.0
        link_primary = 0.0
        link_backup = 0.0
        rx_primary = -40.0
        rx_backup = -40.0
        crc_rate = 2400.0
        incident_open = 1.0
        pending_items = 7.0

    elif scenario == "recover":
        # 恢复过程：链路回来了，业务在爬升
        progress = min(elapsed / 12.0, 1.0)
        availability = 1.0
        connections = _noise(1240 * (0.3 + 0.7 * progress), 0.10)
        p75 = _noise(0.082 + 0.10 * (1 - progress))
        p90 = _noise(0.151 + 0.25 * (1 - progress))
        p99 = _noise(0.342 + 0.50 * (1 - progress))
        rps = _noise(860 * (0.3 + 0.7 * progress), 0.10)
        link_primary = 1.0
        link_backup = 1.0
        rx_primary = _noise(-12.4, 0.03)
        rx_backup = _noise(-12.9, 0.03)
        crc_rate = 2.0
        incident_open = 1.0
        pending_items = 4.0

    with LOCK:
        COUNTERS["crc_primary"] += crc_rate * 2
        COUNTERS["crc_backup"] += crc_rate * 2
        crc_primary = COUNTERS["crc_primary"]
        crc_backup = COUNTERS["crc_backup"]

    return {
        "availability": availability,
        "connections": connections,
        "p75": max(p75, 0.0),
        "p90": max(p90, 0.0),
        "p99": max(p99, 0.0),
        "rps": max(rps, 0.0),
        "link_primary": link_primary,
        "link_backup": link_backup,
        "rx_primary": rx_primary,
        "rx_backup": rx_backup,
        "crc_primary": crc_primary,
        "crc_backup": crc_backup,
        "incident_open": incident_open,
        "pending_items": pending_items,
    }


def render_metrics():
    m = collect()
    lines = [
        "# HELP business_availability 业务可用性（1=可用，0=不可用）",
        "# TYPE business_availability gauge",
        f"business_availability {m['availability']}",
        "# HELP business_connections 活跃连接数",
        "# TYPE business_connections gauge",
        f"business_connections {m['connections']:.0f}",
        "# HELP business_requests_per_second 每秒请求数",
        "# TYPE business_requests_per_second gauge",
        f"business_requests_per_second {m['rps']:.0f}",
        "# HELP business_latency_p75_seconds 请求耗时 P75（秒）",
        "# TYPE business_latency_p75_seconds gauge",
        f"business_latency_p75_seconds {m['p75']:.4f}",
        "# HELP business_latency_p90_seconds 请求耗时 P90（秒）",
        "# TYPE business_latency_p90_seconds gauge",
        f"business_latency_p90_seconds {m['p90']:.4f}",
        "# HELP business_latency_p99_seconds 请求耗时 P99（秒）",
        "# TYPE business_latency_p99_seconds gauge",
        f"business_latency_p99_seconds {m['p99']:.4f}",
        "# HELP link_state 链路状态（1=up，0=down）",
        "# TYPE link_state gauge",
        f'link_state{{link="primary"}} {m["link_primary"]}',
        f'link_state{{link="backup"}} {m["link_backup"]}',
        "# HELP link_optical_rx_dbm 接收光功率（dBm）",
        "# TYPE link_optical_rx_dbm gauge",
        f'link_optical_rx_dbm{{link="primary"}} {m["rx_primary"]:.2f}',
        f'link_optical_rx_dbm{{link="backup"}} {m["rx_backup"]:.2f}',
        "# HELP link_crc_errors_total 接口误码累计数",
        "# TYPE link_crc_errors_total counter",
        f'link_crc_errors_total{{link="primary"}} {m["crc_primary"]:.0f}',
        f'link_crc_errors_total{{link="backup"}} {m["crc_backup"]:.0f}',
        "# HELP incident_open 是否有未闭环事件（1=有）",
        "# TYPE incident_open gauge",
        f"incident_open {m['incident_open']}",
        "# HELP handover_pending_items 交接待跟进事项数",
        "# TYPE handover_pending_items gauge",
        f"handover_pending_items {m['pending_items']}",
        "# HELP exporter_uptime_seconds 本导出器运行时长",
        "# TYPE exporter_uptime_seconds gauge",
        f"exporter_uptime_seconds {time.time() - START:.0f}",
    ]
    return "\n".join(lines) + "\n"


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):  # 静音，避免刷屏
        pass

    def _send(self, code, body, content_type):
        data = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        # 允许本地用 file:// 打开的 preview.html 直接取数（仅演示用途）
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/metrics":
            self._send(200, render_metrics(), "text/plain; version=0.0.4; charset=utf-8")

        elif path == "/control":
            qs = parse_qs(parsed.query)
            scenario = (qs.get("scenario") or ["normal"])[0]
            if scenario not in SCENARIOS:
                self._send(400, f"unknown scenario: {scenario}\n"
                                f"可选: {', '.join(SCENARIOS)}\n", "text/plain; charset=utf-8")
                return
            with LOCK:
                STATE["scenario"] = scenario
                STATE["changed_at"] = time.time()
            self._send(200, f"scenario -> {scenario}\n", "text/plain; charset=utf-8")

        elif path == "/state":
            m = collect()
            payload = {
                "scenario": _current_scenario(),
                "seconds_in_scenario": round(_since_change(), 1),
                "metrics": {k: (round(v, 4) if isinstance(v, float) else v) for k, v in m.items()},
                "hint": "degrade -> down -> recover 可复现一次真实故障的形状",
            }
            self._send(200, json.dumps(payload, ensure_ascii=False, indent=2),
                       "application/json; charset=utf-8")

        elif path == "/":
            body = (
                "值班视图演示用指标生成器\n\n"
                "  /metrics                      Prometheus 抓取端点\n"
                "  /state                        查看当前场景与指标\n"
                "  /control?scenario=degrade     延迟劣化（链路仍 up）\n"
                "  /control?scenario=down        链路中断（主备同时 down）\n"
                "  /control?scenario=recover     恢复\n"
                "  /control?scenario=normal      正常\n"
            )
            self._send(200, body, "text/plain; charset=utf-8")

        else:
            self._send(404, "not found\n", "text/plain; charset=utf-8")


def main():
    ap = argparse.ArgumentParser(description="值班视图演示用指标生成器（全部为模拟数据）")
    ap.add_argument("--port", type=int, default=9100)
    ap.add_argument("--host", default="0.0.0.0")
    ap.add_argument("--scenario", default="normal", choices=SCENARIOS)
    args = ap.parse_args()

    with LOCK:
        STATE["scenario"] = args.scenario
        STATE["changed_at"] = time.time()

    srv = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"[mock-exporter] 监听 http://{args.host}:{args.port}/metrics "
          f"(scenario={args.scenario})")
    print("[mock-exporter] 所有指标均为随机模拟，不含任何真实数据。")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n[mock-exporter] 已停止")


if __name__ == "__main__":
    main()

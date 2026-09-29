#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hostwatch —— 局域网关键主机可用性监控

起因：2026-09-16，192.168.1.111 失联。排查花了十几轮，因为没有任何历史数据：
不知道什么时候断的、断了多久、当时主路由还通不通。最后靠 .111 自己的内核日志
才定位到是 Mesh 子节点 YR1900MG 的桥接卡死，而不是 .111 本身出问题。

这个脚本就是补上那份数据。设计上有两个关键点：

【一】用 TCP 而不是 ICMP 判活
实测 PC ping .111 常态就有 10~26% 丢包（Wi-Fi + 多跳），但同期 TCP 握手 20/20 全过。
拿 ICMP 判活必然误报。所以探测一律走 TCP 端口连接。

【二】必须有"参照主机"
光知道 .111 不通没用 —— 那可能是它自己挂了，也可能是 PC 的 Wi-Fi 断了，
还可能是中间桥接断了。同时探测主路由 M6G 作为参照，三种情况立刻可分：

    目标 down + 参照 up   -> 中间链路/桥接问题（这次就是这种）
    目标 down + 参照 down -> 本机网络问题，与目标无关
    目标 up               -> 正常

本机是 Wi-Fi 直连 M6G 的，Mesh 子节点挂掉时本机不受影响，
所以 PC 是观察这件事的正确位置。

用法：
    python hostwatch.py run                  前台运行
    python hostwatch.py run --quiet          只写日志
    python hostwatch.py status               当前状态 + 最近事件
    python hostwatch.py history --days 7     历史事件
"""

import argparse
import io
import json
import os
import socket
import sys
import time

# ---------------------------------------------------------------- 监控目标
# port 选各自稳定常开的服务：.111 的 sshd、M6G 的 Web 管理页。
# 不用 ICMP，理由见文件头。
TARGETS = [
    # 探 7890（mihomo 代理）而不是 22（sshd）。这不是随意选的 ——
    # 2026-09-18 用 22 端口探了两天，报了 62 次「中断」、累计 57 分钟，
    # 追查 Mesh 子节点追了很久，最后发现全是测量方式自己造的：
    #
    #   .111 -> 主路由     4.3 小时 3845 个样本，失败 0 次
    #   .111 网口 carrier  48 小时 down 过 3 次，而「中断」有 62 次
    #   决定性快照         同一时刻 22 端口不通、同机 7890 通 103ms
    #
    # 原因是拿 TCP 建连即关的方式去探 sshd：sshd 把每次都记成未认证连接
    # （半小时 65 条、一小时 159 条垃圾日志），会触发 PerSourcePenalties
    # （crash:90 / noauth:1）、MaxStartups（10:30:100 概率丢弃）、banner 格式校验
    # 这一堆保护机制。用 shell 的 echo > /dev/tcp 更糟，那个换行符会被 sshd
    # 当成格式错误的 SSH 横幅。
    #
    # 7890 是无状态、无认证的代理端口，探它不会自我污染，
    # 而且它才是真正关心的服务 —— sshd 通不通不代表代理能不能用。
    {"name": "mihomo-box", "host": "192.168.1.111", "port": 7890, "role": "target",
     "note": "代理 / Stash / 小米盒子网关（探 7890，不探 22，理由见上）"},
    {"name": "router-M6G", "host": "192.168.1.11", "port": 80, "role": "reference",
     "note": "主路由，用作参照"},
]

LOG_DIR_DEFAULT = os.path.join(os.environ.get("LOCALAPPDATA", "."), "hostwatch")

# 连续多少轮失败才判定为 down。单轮抖动很常见，尤其经 Wi-Fi；
# 3 轮 × 15 秒 = 45 秒确认窗口，既不误报也不会漏掉真正的长时间中断。
FAIL_THRESHOLD = 3
OK_THRESHOLD = 2          # 恢复确认得快一些，避免把恢复时间记晚


def setup_stdout():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def out(s):
    try:
        print(s)
    except UnicodeEncodeError:
        print(s.encode("unicode_escape").decode())


def probe(host, port, timeout=3.0):
    """TCP 连接探测。返回 (是否通, 毫秒)。"""
    t0 = time.time()
    try:
        s = socket.create_connection((host, port), timeout=timeout)
        s.close()
        return True, int((time.time() - t0) * 1000)
    except Exception:
        return False, None


def log_path(log_dir):
    return os.path.join(log_dir, "events-%s.jsonl" % time.strftime("%Y%m"))


def write_event(log_dir, entry):
    try:
        os.makedirs(log_dir, exist_ok=True)
        with io.open(log_path(log_dir), "a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except OSError:
        pass


def fmt_duration(sec):
    sec = int(sec)
    if sec < 60:
        return "%d 秒" % sec
    if sec < 3600:
        return "%d 分 %d 秒" % (sec // 60, sec % 60)
    return "%d 小时 %d 分" % (sec // 3600, (sec % 3600) // 60)


def diagnose(name, targets_state):
    """把"谁通谁不通"翻译成病因。这是整个脚本的价值所在 ——
    光记录 down 没用，要能一眼看出该去查哪里。

    判定顺序有讲究：先排除「同机另一端口还活着」这种情况，
    否则会把 sshd 拒连误报成链路中断（2026-09-18 就这么误判过，
    追了半天 Mesh 回程，结果 .111 到主路由 503 个样本一次没断）。
    """
    def up_of(n):
        return targets_state.get(n, {}).get("up")

    # 同一台主机的其他端口还通吗？通 = 网络没问题，是这个服务在拒连
    me = next((t for t in TARGETS if t["name"] == name), None)
    if me:
        siblings = [t for t in TARGETS
                    if t["name"] != name and t["host"] == me["host"]]
        alive = [t["name"] for t in siblings if up_of(t["name"])]
        if alive:
            return ("同机的 %s 仍可达 —— 网络正常，是 %d 端口这个服务在拒连"
                    "（sshd 的 PerSourcePenalties 最可能）"
                    % ("/".join(alive), me["port"]))

    ref_up = None
    for t in TARGETS:
        if t["role"] == "reference":
            ref_up = up_of(t["name"])
            break
    if ref_up is None:
        return "无参照主机，无法判断"
    if ref_up:
        return "主路由仍可达，且本机所有端口都不通 —— 中间链路问题（Mesh 桥接 / 交换机 / 网线）"
    return "主路由也不可达 —— 本机网络问题（Wi-Fi 掉线等），与目标主机无关"


def cmd_run(args):
    log_dir = args.log or LOG_DIR_DEFAULT
    os.makedirs(log_dir, exist_ok=True)

    state = {}
    for t in TARGETS:
        state[t["name"]] = {"up": None, "fail_streak": 0, "ok_streak": 0,
                            "since": time.time(), "down_at": None}

    if not args.quiet:
        out("hostwatch 已启动")
        for t in TARGETS:
            out("  %-12s %s:%d  (%s)" % (t["name"], t["host"], t["port"], t["note"]))
        out("  日志 %s" % log_dir)
        out("  每 %d 秒探测一轮，连续 %d 轮失败才判定中断" % (args.interval, FAIL_THRESHOLD))
        out("  Ctrl+C 退出\n")

    while True:
        now = time.time()
        snapshot = {}

        for t in TARGETS:
            st = state[t["name"]]
            up, ms = probe(t["host"], t["port"], timeout=args.timeout)

            if up:
                st["ok_streak"] += 1
                st["fail_streak"] = 0
            else:
                st["fail_streak"] += 1
                st["ok_streak"] = 0

            # 首次探测直接定状态，不走阈值（否则启动时会误报一次恢复）
            if st["up"] is None:
                st["up"] = up
                st["since"] = now
            snapshot[t["name"]] = {"up": st["up"], "ms": ms}

        # 判定状态翻转。先收集完所有主机的本轮结果再判，
        # 这样 diagnose() 拿到的是同一时刻的快照，不会前后错位。
        for t in TARGETS:
            st = state[t["name"]]
            name = t["name"]

            if st["up"] and st["fail_streak"] >= FAIL_THRESHOLD:
                st["up"] = False
                st["down_at"] = now
                reason = diagnose(name, snapshot)
                entry = {"ts": time.strftime("%Y-%m-%d %H:%M:%S"), "event": "down",
                         "name": name, "host": t["host"], "diagnosis": reason,
                         "peers": {k: v["up"] for k, v in snapshot.items()}}
                write_event(log_dir, entry)
                if not args.quiet:
                    out("[%s] 中断  %s (%s)" % (entry["ts"], name, t["host"]))
                    out("    %s\n" % reason)

            elif (st["up"] is False) and st["ok_streak"] >= OK_THRESHOLD:
                st["up"] = True
                dur = now - (st["down_at"] or now)
                entry = {"ts": time.strftime("%Y-%m-%d %H:%M:%S"), "event": "up",
                         "name": name, "host": t["host"],
                         "outage_seconds": int(dur)}
                write_event(log_dir, entry)
                if not args.quiet:
                    out("[%s] 恢复  %s —— 中断持续 %s\n"
                        % (entry["ts"], name, fmt_duration(dur)))
                st["down_at"] = None

        time.sleep(args.interval)


def load_events(log_dir, months=2):
    if not os.path.isdir(log_dir):
        return []
    files = sorted(fn for fn in os.listdir(log_dir)
                   if fn.startswith("events-") and fn.endswith(".jsonl"))
    rows = []
    for fn in files[-months:]:
        try:
            for line in io.open(os.path.join(log_dir, fn), encoding="utf-8"):
                line = line.strip()
                if line:
                    rows.append(json.loads(line))
        except (OSError, ValueError):
            continue
    return rows


def cmd_status(args):
    log_dir = args.log or LOG_DIR_DEFAULT
    out("当前探测结果：")
    for t in TARGETS:
        up, ms = probe(t["host"], t["port"], timeout=3.0)
        out("  %-12s %-16s %s%s" % (
            t["name"], "%s:%d" % (t["host"], t["port"]),
            "通" if up else "不通",
            ("  %dms" % ms) if ms is not None else ""))

    rows = load_events(log_dir)
    if not rows:
        out("\n还没有记录到任何中断事件。")
        return
    out("\n最近 %d 条事件：" % min(args.limit, len(rows)))
    for e in rows[-args.limit:]:
        if e["event"] == "down":
            out("  %s  中断  %s" % (e["ts"], e["name"]))
            out("      %s" % e.get("diagnosis", ""))
        else:
            out("  %s  恢复  %s（持续 %s）"
                % (e["ts"], e["name"], fmt_duration(e.get("outage_seconds", 0))))


def cmd_history(args):
    log_dir = args.log or LOG_DIR_DEFAULT
    rows = load_events(log_dir)
    cutoff = time.time() - args.days * 86400
    kept = []
    for e in rows:
        try:
            ts = time.mktime(time.strptime(e["ts"], "%Y-%m-%d %H:%M:%S"))
        except ValueError:
            continue
        if ts >= cutoff:
            kept.append(e)
    if not kept:
        out("最近 %d 天没有中断记录。" % args.days)
        return

    # 按主机汇总，给出总中断次数和累计时长 —— 判断某台设备是不是反复出问题
    stats = {}
    for e in kept:
        s = stats.setdefault(e["name"], {"count": 0, "total": 0})
        if e["event"] == "down":
            s["count"] += 1
        else:
            s["total"] += e.get("outage_seconds", 0)

    out("最近 %d 天汇总：" % args.days)
    for name, s in stats.items():
        out("  %-12s 中断 %d 次，累计 %s" % (name, s["count"], fmt_duration(s["total"])))
    out("")
    for e in kept:
        if e["event"] == "down":
            out("  %s  中断  %-12s %s" % (e["ts"], e["name"], e.get("diagnosis", "")))
        else:
            out("  %s  恢复  %-12s 持续 %s"
                % (e["ts"], e["name"], fmt_duration(e.get("outage_seconds", 0))))


def main():
    ap = argparse.ArgumentParser(description="局域网关键主机可用性监控")
    sub = ap.add_subparsers(dest="cmd", required=True)

    r = sub.add_parser("run", help="持续监控")
    r.add_argument("--interval", type=int, default=15, help="探测间隔秒，默认 15")
    r.add_argument("--timeout", type=float, default=3.0, help="单次探测超时秒")
    r.add_argument("--log", help="日志目录")
    r.add_argument("--quiet", action="store_true", help="只写日志不打印")
    r.set_defaults(func=cmd_run)

    s = sub.add_parser("status", help="当前状态 + 最近事件")
    s.add_argument("--log")
    s.add_argument("--limit", type=int, default=10)
    s.set_defaults(func=cmd_status)

    h = sub.add_parser("history", help="历史中断汇总")
    h.add_argument("--log")
    h.add_argument("--days", type=int, default=7)
    h.set_defaults(func=cmd_history)

    args = ap.parse_args()
    setup_stdout()
    args.func(args)


if __name__ == "__main__":
    main()

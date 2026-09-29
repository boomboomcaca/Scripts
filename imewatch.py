#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
imewatch —— Win11 中文输入法候选条消失的监控与自动恢复

起因：2026-09-19，中文输入法突然不显示候选词。查下来发现 Win11 的拼音候选条
不是独立窗口，它画在 TextInputHost.exe 那个全屏透明的 "Windows 输入体验" 窗口里。
故障时这个窗口沉到了 z 序最底下 —— 候选条照常在绘制，但被所有应用窗口盖住，
一个字也看不见。重启 TextInputHost / ctfmon / ChsIME 就能恢复。

麻烦的是根因不明：事件日志里这三个进程一条错误记录都没有，它是静默沉底的。
所以这个脚本有两个目的，而且抓现场比自动恢复更重要：

【一】抓现场
一旦检测到沉底，先把当时的完整 z 序、前台窗口、所有全屏覆盖层进程写进日志，
然后才动手恢复。攒上几次就能看出是谁在捣鬼 —— 本机嫌疑不少：PixPin、RedDot、
RzSynapse、成都银行的键盘过滤驱动(gvinput.sys/gvinputmf.sys)，还有 Claude
做计算机控制时挂的 cua-driver 全屏覆盖层。

【二】自动恢复
检测到就按顺序重启那三个进程，平时基本感觉不到这个故障。

判据（来自 2026-09-19 那次实测的两组数据，只有两个样本，属启发式）：
    故障时  输入宿主 z13/共15，下面只剩 Program Manager
    正常时  输入宿主 z6/共19，下面还压着好几个在屏窗口
看的是"它下面还剩不剩在屏窗口"，而不是"上面压了几个" —— 上面压几个正常时也有，
我第一版就是按上面数的，把健康状态误报成了故障。

注意：脚本无法自测是否修好。微软拼音对程序合成的按键不响应（防键盘记录的设计），
候选条只有真实按键才会弹出来，所以最终验证只能靠人敲一下。

日志用本机本地时间。本机时区是美国东部，比局域网其它设备慢 12 小时，
跟 .111 之类的日志对照时记得换算。

用法：
    python imewatch.py check              一次性诊断
    python imewatch.py fix                一次性恢复
    python imewatch.py run                前台常驻
    python imewatch.py run --quiet        常驻，只写日志（配 StartImewatch.vbs）
    python imewatch.py status             当前状态 + 实时诊断
    python imewatch.py history --days 7   历史事件
"""

import argparse
import ctypes
import json
import os
import subprocess
import sys
import time
from ctypes import wintypes

# Windows 控制台默认 cp1252，直接 print 中文会抛 UnicodeEncodeError
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ---------------------------------------------------------------- 常量

# 要重启的三个进程，顺序有讲究：先重启画候选条的宿主，再重启文本服务框架，
# 最后才是输入法本身，这样 ChsIME 起来时对接到的是已经就绪的新 ctfmon。
IME_PROCS = ["TextInputHost.exe", "ctfmon.exe", "ChsIME.exe"]

# ctfmon 被杀后不一定会自动拉起，而它没了输入法就彻底不工作，必须能补一刀
CTFMON_PATH = os.path.join(
    os.environ.get("WINDIR", r"C:\Windows"), "System32", "ctfmon.exe")

DATA_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "imewatch")
STATE_FILE = os.path.join(DATA_DIR, "state.json")
PID_FILE = os.path.join(DATA_DIR, "imewatch.pid")

DEFAULT_INTERVAL = 30       # 秒
DEFAULT_COOLDOWN = 180      # 修完多久内不再动手，避免反复重启
HEARTBEAT_SECONDS = 3600    # 一切正常时每小时记一行，证明自己还活着

# ---------------------------------------------------------------- Win32

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
user32.EnumWindows.argtypes = [WNDENUMPROC, wintypes.LPARAM]
user32.EnumWindows.restype = wintypes.BOOL
user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
user32.IsWindowVisible.argtypes = [wintypes.HWND]
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowThreadProcessId.argtypes = [
    wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int]
user32.GetForegroundWindow.restype = wintypes.HWND

GWL_EXSTYLE = -20
WS_EX_TOPMOST = 0x8
WS_EX_LAYERED = 0x80000
WS_EX_TRANSPARENT = 0x20
WS_EX_TOOLWINDOW = 0x80

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.QueryFullProcessImageNameW.argtypes = [
    wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR,
    ctypes.POINTER(wintypes.DWORD)]
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

try:
    # 不声明 DPI 感知的话，225% 缩放下拿到的是虚拟化坐标，
    # 尺寸阈值和"最小化"判断会跟着飘
    ctypes.WinDLL("shcore").SetProcessDpiAwareness(2)
except Exception:
    pass

_proc_name_cache = {}


def proc_name(pid):
    if pid in _proc_name_cache:
        return _proc_name_cache[pid]
    name = "?"
    h = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if h:
        buf = ctypes.create_unicode_buffer(4096)
        size = wintypes.DWORD(4096)
        if kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
            name = os.path.basename(buf.value)
        kernel32.CloseHandle(h)
    _proc_name_cache[pid] = name
    return name


class Win:
    __slots__ = ("z", "hwnd", "pid", "proc", "title", "cls",
                 "l", "t", "w", "h", "ex")

    @property
    def minimized(self):
        return self.l < -30000

    @property
    def is_desktop(self):
        # 桌面本体永远在最底下，不能算"压在输入宿主下面的窗口"
        return self.cls in ("Progman", "WorkerW")

    @property
    def on_screen(self):
        return (not self.minimized) and self.w > 200 and self.h > 200

    @property
    def flags(self):
        f = []
        if self.ex & WS_EX_TOPMOST:
            f.append("TOPMOST")
        if self.ex & WS_EX_LAYERED:
            f.append("LAYERED")
        if self.ex & WS_EX_TRANSPARENT:
            f.append("TRANSPARENT")
        if self.ex & WS_EX_TOOLWINDOW:
            f.append("TOOL")
        return ",".join(f)

    def line(self):
        return "z%-3d %-22s %6d,%-6d %5dx%-5d %-26s '%s'" % (
            self.z, self.proc, self.l, self.t, self.w, self.h,
            self.flags, self.title)


def walk_windows():
    """按 z 序返回所有可见且非零尺寸的窗口。

    EnumWindows 的回调顺序本身就是 z 序，最前面的先来。
    """
    out = []
    counter = [0]

    def cb(hwnd, _):
        if not user32.IsWindowVisible(hwnd):
            return True
        r = wintypes.RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(r)):
            return True
        w, h = r.right - r.left, r.bottom - r.top
        if w <= 1 or h <= 1:
            return True
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        tbuf = ctypes.create_unicode_buffer(256)
        user32.GetWindowTextW(hwnd, tbuf, 256)
        cbuf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, cbuf, 256)
        win = Win()
        win.z = counter[0]
        counter[0] += 1
        win.hwnd = hwnd
        win.pid = pid.value
        win.proc = proc_name(pid.value)
        win.title = tbuf.value
        win.cls = cbuf.value
        win.l, win.t, win.w, win.h = r.left, r.top, w, h
        win.ex = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        out.append(win)
        return True

    user32.EnumWindows(WNDENUMPROC(cb), 0)
    return out


# ---------------------------------------------------------------- 判定

class State:
    def __init__(self, wins):
        self.wins = wins
        self.host = next(
            (w for w in wins
             if w.proc.lower() == "textinputhost.exe" and not w.minimized),
            None)
        if self.host is None:
            self.below = []
        else:
            self.below = [w for w in wins
                          if w.z > self.host.z and w.on_screen
                          and not w.is_desktop]

    @property
    def idle(self):
        """输入宿主没有在屏窗口 —— 闲置，此时看不出健康与否"""
        return self.host is None

    @property
    def sunk(self):
        """沉底：它下面除了桌面什么都不剩了"""
        return (self.host is not None) and len(self.below) == 0

    def summary(self):
        if self.idle:
            return "输入宿主闲置（无在屏窗口），看不出问题"
        return "输入宿主 z%d/共%d，下面还有 %d 个在屏窗口 -> %s" % (
            self.host.z, len(self.wins), len(self.below),
            "沉底(疑似故障)" if self.sunk else "正常")


def foreground_info():
    hwnd = user32.GetForegroundWindow()
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    tbuf = ctypes.create_unicode_buffer(256)
    user32.GetWindowTextW(hwnd, tbuf, 256)
    return "%s '%s'" % (proc_name(pid.value), tbuf.value)


def overlay_suspects(wins):
    """全屏的分层/透明/置顶窗口 —— 最可能干扰 z 序的那一类"""
    return [w for w in wins
            if w.w >= 1920 and w.h >= 1080
            and (w.ex & (WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOPMOST))]


# ---------------------------------------------------------------- 进程操作

def running_procs():
    """用 tasklist 拿 进程名(小写) -> [pid]，避免引入额外依赖"""
    try:
        out = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=20).stdout
    except Exception:
        return {}
    res = {}
    for line in out.splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) >= 2:
            name = parts[0].strip('"').lower()
            try:
                res.setdefault(name, []).append(int(parts[1]))
            except ValueError:
                pass
    return res


def restart_ime(log):
    """按 IME_PROCS 的顺序重启三个进程"""
    for name in IME_PROCS:
        before = running_procs().get(name.lower(), [])
        if not before:
            log("%s 本来就没在跑，跳过" % name)
            continue
        subprocess.run(["taskkill", "/F", "/IM", name],
                       capture_output=True, text=True, timeout=20)
        time.sleep(2)
        after = running_procs().get(name.lower(), [])
        if not after and name.lower() == "ctfmon.exe":
            try:
                subprocess.Popen([CTFMON_PATH])
                time.sleep(2)
                after = running_procs().get(name.lower(), [])
            except Exception as e:
                log("ctfmon 手动拉起失败: %s" % e)
        if after:
            log("%s: %s -> %s" % (name, before, after))
        else:
            log("%s: %s -> 已结束，等待按需拉起" % (name, before))


# ---------------------------------------------------------------- 日志

def ensure_dir():
    os.makedirs(DATA_DIR, exist_ok=True)


def log_path():
    return os.path.join(DATA_DIR, "imewatch-%s.log" % time.strftime("%Y-%m"))


def make_logger(quiet):
    ensure_dir()

    def log(msg, to_file=True):
        line = "[%s] %s" % (time.strftime("%Y-%m-%d %H:%M:%S"), msg)
        if not quiet:
            print(line, flush=True)
        if to_file:
            try:
                with open(log_path(), "a", encoding="utf-8") as f:
                    f.write(line + "\n")
            except Exception:
                pass
    return log


def save_state(d):
    ensure_dir()
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def load_state():
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


# ---------------------------------------------------------------- 单实例

def claim_pidfile(log):
    """防止起两份。

    注意不能用"命令行里含 imewatch.py"去找同类进程 —— 那个模式会匹配到自己，
    是踩过的坑（2026-09-17 用 pkill -f 把自己的 shell 干掉过）。用 pid 文件。
    """
    ensure_dir()
    if os.path.exists(PID_FILE):
        try:
            with open(PID_FILE, encoding="utf-8") as f:
                old = int(f.read().strip())
        except Exception:
            old = None
        if old and old != os.getpid():
            alive = any(old in pids for pids in running_procs().values())
            if alive:
                log("已有一份 imewatch 在跑（PID %d），本次退出" % old,
                    to_file=False)
                return False
    with open(PID_FILE, "w", encoding="utf-8") as f:
        f.write(str(os.getpid()))
    return True


def release_pidfile():
    try:
        if os.path.exists(PID_FILE):
            with open(PID_FILE, encoding="utf-8") as f:
                if f.read().strip() == str(os.getpid()):
                    os.remove(PID_FILE)
    except Exception:
        pass


# ---------------------------------------------------------------- 命令

def cmd_check(args):
    wins = walk_windows()
    st = State(wins)
    print("=== 输入法状态诊断 ===")
    names = running_procs()
    for p in IME_PROCS:
        pids = names.get(p.lower(), [])
        print("  %-18s %s" % (
            p, ("PID " + ", ".join(map(str, pids))) if pids
            else "未运行（按需启动）"))
    print()
    print("  " + st.summary())
    if st.host is None:
        print("  在任意输入框敲一个拼音让它现身，然后再跑一次。")
        return 0
    print("  输入宿主: " + st.host.line())
    if st.sunk:
        print()
        print("  >>> 疑似沉底。恢复： python imewatch.py fix")
        print("  当时的 z 序：")
        for w in wins:
            print("    " + w.line())
    else:
        print("  下面的在屏窗口（取前 4 个）：")
        for w in st.below[:4]:
            print("    " + w.line())
        print()
        print("  若候选词仍不出，那是别的原因。优先怀疑带键盘过滤驱动的安全输入")
        print("  控件（网银那类，如 BOCDSecInput.exe + gvinput.sys），"
              "先结束其用户态进程试试。")
    return 0


def cmd_fix(args):
    log = make_logger(quiet=False)
    log("手动恢复。恢复前：" + State(walk_windows()).summary())
    restart_ime(log)
    time.sleep(1)
    log("恢复后：" + State(walk_windows()).summary())
    print()
    print("现在敲一个拼音验证 —— 脚本测不了，微软拼音不响应合成按键。")
    return 0


def snapshot_lines(wins, st):
    out = ["前台窗口: " + foreground_info(),
           "输入宿主: " + (st.host.line() if st.host else "(无在屏窗口)"),
           "全屏覆盖层嫌疑："]
    sus = overlay_suspects(wins)
    out += ["    " + w.line() for w in sus] or ["    (无)"]
    out.append("完整 z 序：")
    out += ["    " + w.line() for w in wins]
    return out


def cmd_run(args):
    log = make_logger(args.quiet)
    if not claim_pidfile(log):
        return 1
    log("imewatch 启动，间隔 %ds，冷却 %ds，日志 %s"
        % (args.interval, args.cooldown, log_path()))

    last_sunk = False
    last_fix = 0.0
    last_beat = time.time()
    fixes = 0

    try:
        while True:
            wins = walk_windows()
            st = State(wins)
            now = time.time()

            if st.sunk and not last_sunk:
                # 抓现场永远排在恢复前面 —— 这份快照才是找根因的唯一线索
                log("!! 检测到输入宿主沉底，候选条会被所有窗口盖住")
                for line in snapshot_lines(wins, st):
                    log("   " + line)
                if now - last_fix < args.cooldown:
                    log("   距上次恢复不足 %ds，本次不动手" % args.cooldown)
                else:
                    fixes += 1
                    log("   开始第 %d 次自动恢复" % fixes)
                    restart_ime(lambda m: log("   " + m))
                    last_fix = time.time()
                    time.sleep(1)
                    log("   恢复后：" + State(walk_windows()).summary())
                    if fixes >= 5:
                        log("   已恢复 %d 次，触发源显然还在。"
                            "翻本日志里历次 z 序快照找共同点。" % fixes)
            elif last_sunk and not st.sunk:
                log("输入宿主已回到正常位置")

            last_sunk = st.sunk

            if now - last_beat >= HEARTBEAT_SECONDS:
                log("心跳：" + st.summary())
                last_beat = now

            save_state({
                "updated": time.strftime("%Y-%m-%d %H:%M:%S"),
                "summary": st.summary(),
                "sunk": st.sunk,
                "idle": st.idle,
                "fixes": fixes,
                "pid": os.getpid(),
            })
            time.sleep(args.interval)
    except KeyboardInterrupt:
        log("收到中断，退出")
    finally:
        release_pidfile()
    return 0


def cmd_status(args):
    s = load_state()
    if not s:
        print("还没有状态记录，说明 imewatch run 没跑过。")
    else:
        print("最后更新: %s" % s.get("updated"))
        print("当时状态: %s" % s.get("summary"))
        print("已自动恢复 %s 次，守护进程 PID %s"
              % (s.get("fixes"), s.get("pid")))
    print()
    print("实时诊断：")
    return cmd_check(args)


def cmd_history(args):
    if not os.path.isdir(DATA_DIR):
        print("还没有日志目录，说明 imewatch run 没跑过。")
        return 0
    cutoff = time.time() - args.days * 86400
    shown = 0
    for fn in sorted(os.listdir(DATA_DIR)):
        if not (fn.startswith("imewatch-") and fn.endswith(".log")):
            continue
        with open(os.path.join(DATA_DIR, fn),
                  encoding="utf-8", errors="replace") as f:
            for line in f:
                if not line.startswith("["):
                    continue
                try:
                    ts = time.mktime(
                        time.strptime(line[1:20], "%Y-%m-%d %H:%M:%S"))
                except ValueError:
                    continue
                if ts < cutoff:
                    continue
                # 只列事件本身，跳过心跳和快照明细（明细缩进了 3 个空格）
                if "心跳" in line or line[22:].startswith("   "):
                    continue
                print(line.rstrip())
                shown += 1
    if not shown:
        print("最近 %d 天没有事件记录。" % args.days)
    return 0


def main():
    ap = argparse.ArgumentParser(
        description="Win11 输入法候选条沉底的监控与恢复")
    sub = ap.add_subparsers(dest="cmd")

    sub.add_parser("check", help="一次性诊断")
    sub.add_parser("fix", help="一次性恢复")

    r = sub.add_parser("run", help="常驻监控")
    r.add_argument("--quiet", action="store_true", help="只写日志不打印")
    r.add_argument("--interval", type=int, default=DEFAULT_INTERVAL,
                   help="检查间隔秒数")
    r.add_argument("--cooldown", type=int, default=DEFAULT_COOLDOWN,
                   help="两次恢复的最小间隔秒数")

    sub.add_parser("status", help="当前状态 + 实时诊断")

    h = sub.add_parser("history", help="历史事件")
    h.add_argument("--days", type=int, default=7)

    args = ap.parse_args()
    if args.cmd == "check":
        return cmd_check(args)
    if args.cmd == "fix":
        return cmd_fix(args)
    if args.cmd == "run":
        return cmd_run(args)
    if args.cmd == "status":
        return cmd_status(args)
    if args.cmd == "history":
        return cmd_history(args)
    ap.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())

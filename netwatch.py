#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
netwatch —— Clash / mihomo 路由诊断

起因：同一个坑踩了三次。主域名规则配好了，但页面的图片、视频、静态资源
在另一个域名上（ttcache.com / pvvstream.pro / sb-cd.com），漏掉就是缩略图
空白或者裸 HTML。人眼防不住，机器一眼就看出来。

用法：
    python netwatch.py site spankbang.com
    python netwatch.py site taobao.com --watch 15
    python netwatch.py check assets.sb-cd.com
    python netwatch.py rules sb-cd.com

设计要点见 README 段落，核心是把失败拆成四层分别计时：
DNS / 隧道 / TLS / HTTP —— 不同层失败对应完全不同的病因。
"""

import argparse
import io
import json
import os
import re
import socket
import ssl
import sys
import time
import urllib.error
import urllib.request

# UA 的版本号必须是真实存在的。2026-09-19 踩过：原来写的 Chrome/152.0.0.0 根本不存在，
# Cloudflare 直接回 403，于是本工具把 pornoxo.com 误判成"被质询挡住"，实际带 Chrome/131 一探就是 200。
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")

VERGE_DIR = os.path.join(
    os.environ.get("APPDATA", ""),
    "io.github.clash-verge-rev.clash-verge-rev")


# ---------------------------------------------------------------- 输出编码
# Windows 控制台默认 cp1252，组名里的 ♻️自动选择 会直接让脚本崩掉。
# 这个坑在排查过程中打断过四次，所以开头就按死。
def setup_stdout(ascii_mode):
    if ascii_mode:
        return
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def out(s):
    try:
        print(s)
    except UnicodeEncodeError:
        print(s.encode("unicode_escape").decode())


# ---------------------------------------------------------------- 配置发现
def read_runtime_config():
    """从 clash-verge.yaml 里抠出控制器地址、secret、混合端口。

    故意不引 yaml 库：这台机器上跑的是标准库 Python，少一个依赖少一份麻烦。
    这几个字段格式固定，正则足够。
    """
    path = os.path.join(VERGE_DIR, "clash-verge.yaml")
    if not os.path.exists(path):
        raise SystemExit("找不到 %s" % path)
    text = io.open(path, encoding="utf-8", errors="replace").read()

    def grab(key, default=None):
        m = re.search(r"^%s:\s*(.+?)\s*$" % re.escape(key), text, re.M)
        return m.group(1).strip().strip("'\"") if m else default

    ctrl = grab("external-controller", "127.0.0.1:9097")
    secret = grab("secret", "")
    port = grab("mixed-port") or grab("port") or "7890"
    return ctrl, secret, int(port)


CTRL, SECRET, MIXED_PORT = None, None, None


def api(path, method="GET", payload=None, timeout=20):
    url = "http://%s%s" % (CTRL, path)
    headers = {"Accept": "application/json"}
    if SECRET:
        headers["Authorization"] = "Bearer " + SECRET
    data = None
    if payload is not None:
        data = json.dumps(payload).encode()
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        body = r.read()
    return json.loads(body) if body else {}


# ---------------------------------------------------------------- 规则匹配
# mihomo 是自上而下第一条命中生效，不按具体程度排序 —— 顺序就是优先级。
# GEOSITE / RULE-SET / GEOIP 这类规则本地判不了（要外部数据库），
# 遇到就如实标成"无法判定"，而不是假装匹配成功。宁可说不知道，不要给错答案。
UNDECIDABLE = {"GeoSite", "RuleSet", "GeoIP", "IPCIDR", "IPCIDR6",
               "SrcIPCIDR", "DstPort", "SrcPort", "ProcessName", "ProcessPath",
               "InType", "Network", "Uid"}


def load_rules():
    return api("/rules").get("rules", [])


def match_domain(host, rules):
    """返回 (命中项, 该命中之前有多少条判不了的规则, 被遮蔽的同名规则列表)"""
    host = host.lower().lstrip(".")
    undecidable_before = 0
    hit = None
    shadowed = []
    for i, r in enumerate(rules):
        typ = r.get("type", "")
        payload = (r.get("payload") or "").lower()
        ok = False
        if typ == "Domain":
            ok = host == payload
        elif typ == "DomainSuffix":
            ok = host == payload or host.endswith("." + payload)
        elif typ == "DomainKeyword":
            ok = payload in host
        elif typ == "Match":
            ok = True
        elif typ in UNDECIDABLE:
            if hit is None:
                undecidable_before += 1
            continue
        if ok:
            if hit is None:
                hit = (i, r)
            else:
                shadowed.append((i, r))
    return hit, undecidable_before, shadowed


# ---------------------------------------------------------------- 分层探测
def probe(host, port=443, timeout=20):
    """分四层测，每层单独计时。这是整个工具最有价值的部分。

    TLS 无响应      -> 出口 IP 被目标封了，换节点
    TLS 通但 403    -> 机器人质询，浏览器能过，别改配置
    全通但 TLS 慢   -> 解析到远端节点，查 DNS
    全通且快        -> 问题不在网络层，去查浏览器插件
    """
    res = {"host": host, "tunnel_ms": None, "tls_ms": None,
           "status": None, "server": None, "cf": None, "error": None}
    t0 = time.time()
    sock = None
    try:
        sock = socket.create_connection(("127.0.0.1", MIXED_PORT), timeout=timeout)
        req = ("CONNECT %s:%d HTTP/1.1\r\nHost: %s:%d\r\n"
               "User-Agent: %s\r\n\r\n" % (host, port, host, port, UA))
        sock.sendall(req.encode())
        buf = b""
        while b"\r\n\r\n" not in buf:
            chunk = sock.recv(4096)
            if not chunk:
                break
            buf += chunk
        res["tunnel_ms"] = int((time.time() - t0) * 1000)
        first = buf.split(b"\r\n", 1)[0].decode("latin-1")
        if " 200" not in first:
            res["error"] = "隧道失败: " + first
            return res

        t1 = time.time()
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE   # 只测连通性，不做证书校验
        try:
            tls = ctx.wrap_socket(sock, server_hostname=host)
        except (ssl.SSLError, socket.timeout, OSError) as e:
            res["tls_ms"] = int((time.time() - t1) * 1000)
            res["error"] = "TLS 握手失败: %s" % (e.__class__.__name__)
            return res
        res["tls_ms"] = int((time.time() - t1) * 1000)

        tls.settimeout(timeout)
        get = ("GET / HTTP/1.1\r\nHost: %s\r\nUser-Agent: %s\r\n"
               "Accept: */*\r\nConnection: close\r\n\r\n" % (host, UA))
        tls.sendall(get.encode())
        head = b""
        while b"\r\n\r\n" not in head and len(head) < 65536:
            try:
                chunk = tls.recv(4096)
            except (socket.timeout, ssl.SSLError):
                break
            if not chunk:
                break
            head += chunk
        text = head.decode("latin-1", "replace")
        m = re.match(r"HTTP/[\d.]+ (\d{3})", text)
        if m:
            res["status"] = int(m.group(1))
        for line in text.split("\r\n"):
            low = line.lower()
            if low.startswith("server:"):
                res["server"] = line.split(":", 1)[1].strip()
            elif low.startswith("cf-mitigated:"):
                res["cf"] = line.split(":", 1)[1].strip()
        try:
            tls.close()
        except Exception:
            pass
        sock = None
    except Exception as e:
        res["error"] = "%s: %s" % (e.__class__.__name__, e)
    finally:
        if sock is not None:
            try:
                sock.close()
            except Exception:
                pass
    return res


def verdict(p):
    """把探测结果翻译成人话 —— 工具的意义在于给结论，不是给数字"""
    if p["error"] and "隧道" in p["error"]:
        return "代理隧道就没建起来，检查 mihomo 本身"
    if p["error"] and "TLS" in p["error"]:
        return "TLS 握手死掉：出口 IP 大概率被目标站封了，换节点"
    if p["error"]:
        return p["error"]
    if p["cf"] == "challenge" or (p["status"] == 403 and p["server"] == "cloudflare"):
        return ("Cloudflare 403：可能是真质询，也可能是本工具的请求特征被识破 —— "
                "换真实浏览器打开确认后再下结论")
    if p["status"] and 200 <= p["status"] < 400:
        if p["tls_ms"] and p["tls_ms"] > 1000:
            return "通了但 TLS %dms 偏慢：查 DNS 是不是解析到了远端节点" % p["tls_ms"]
        return "正常"
    return "HTTP %s" % p["status"]


# ---------------------------------------------------------------- 域名发现
def fetch(url, timeout=25):
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Accept": "text/html,application/xhtml+xml,*/*",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "identity",
    })
    proxy = urllib.request.ProxyHandler({
        "http": "http://127.0.0.1:%d" % MIXED_PORT,
        "https": "http://127.0.0.1:%d" % MIXED_PORT})
    opener = urllib.request.build_opener(proxy)
    with opener.open(req, timeout=timeout) as r:
        return r.read().decode("utf-8", "replace"), r.status


HOST_RE = re.compile(r"(?:https?:)?//([a-zA-Z0-9][a-zA-Z0-9.-]{1,80}\.[a-z]{2,12})")
IP_RE = re.compile(r"^\d{1,3}(\.\d{1,3}){3}$|^\[?[0-9a-fA-F:]+\]?$")


def extract_hosts(html):
    """抽出页面引用的所有主机名，按出现次数排序。

    先把 JSON 里转义的 \\/ 还原成 / —— 否则播放地址那种埋在 JS 变量里的
    域名全会漏掉，spankbang 那次就差点栽在这。
    """
    plain = html.replace(chr(92) + "/", "/")
    counts = {}
    for h in HOST_RE.findall(plain):
        h = h.lower().rstrip(".")
        counts[h] = counts.get(h, 0) + 1
    return sorted(counts.items(), key=lambda kv: -kv[1])


def registrable(host):
    """粗略取主域名，用于判断是不是站点自己的域。
    没有 PSL，二级后缀（.com.cn 等）会判不准，仅用于分组提示。
    """
    parts = host.split(".")
    if len(parts) >= 3 and parts[-2] in ("com", "net", "org", "co", "gov", "edu"):
        return ".".join(parts[-3:])
    return ".".join(parts[-2:])


# ---------------------------------------------------------------- 实时观察
def watch(seconds, interval=0.4):
    """轮询 /connections 收集域名。

    刻意不用 WebSocket：标准库没有 ws 客户端，为这个引依赖不值。
    更重要的是 —— 快照式查询有个致命陷阱：连接失败的域名根本不会出现在
    列表里。排查 spankbang 时我就是据此误判"没有独立 CDN"，结论是错的。
    所以这里只作为补充信号，真正的判断靠 probe()。
    """
    seen = {}
    end = time.time() + seconds
    while time.time() < end:
        try:
            conns = api("/connections", timeout=8).get("connections") or []
        except Exception:
            time.sleep(interval)
            continue
        for c in conns:
            meta = c.get("metadata") or {}
            h = meta.get("host") or meta.get("destinationIP") or ""
            if not h:
                continue
            chain = " <- ".join(c.get("chains") or [])
            seen[h] = {"rule": c.get("rule"), "payload": c.get("rulePayload"),
                       "chain": chain, "ip": meta.get("destinationIP", "")}
        time.sleep(interval)
    return seen


# ---------------------------------------------------------------- DNS 对比
def dns_via_mihomo(name):
    try:
        d = api("/dns/query?name=%s&type=A" % name, timeout=15)
        return [a["data"] for a in (d.get("Answer") or []) if a.get("type") == 1]
    except Exception:
        return []


def dns_via_alidns(name):
    """从本机出口问阿里 DNS。和 mihomo 的结果不一致，往往意味着
    国内域名被解析到了境外 CDN 节点 —— 淘宝慢 20 倍就是这么来的。"""
    try:
        html, _ = fetch("https://dns.alidns.com/resolve?name=%s&type=A" % name)
        d = json.loads(html)
        return [a["data"] for a in (d.get("Answer") or []) if a.get("type") == 1]
    except Exception:
        return []


# ---------------------------------------------------------------- 命令
def show_rule(host, rules, indent="  "):
    hit, undecidable, shadowed = match_domain(host, rules)
    if not hit:
        out(indent + "无命中（异常，配置里应该有 MATCH 兜底）")
        return None
    i, r = hit
    tag = "  <- 兜底" if r.get("type") == "Match" else ""
    out("%s#%-5d %-13s %-28s -> %s%s" %
        (indent, i, r.get("type"), r.get("payload") or "-", r.get("proxy"), tag))
    if undecidable:
        out(indent + "注意：命中之前有 %d 条 GEOSITE/RULE-SET/GEOIP 规则本地判不了，"
                     "实际可能更早被它们截走" % undecidable)
    for j, s in shadowed[:3]:
        out("%s被遮蔽: #%-5d %-13s %-24s -> %s" %
            (indent, j, s.get("type"), s.get("payload"), s.get("proxy")))
    return r


def cmd_rules(args):
    rules = load_rules()
    out("规则总数 %d" % len(rules))
    for host in args.domains:
        out("\n[%s]" % host)
        show_rule(host, rules)


def cmd_check(args):
    rules = load_rules()
    for host in args.domains:
        out("\n[%s]" % host)
        show_rule(host, rules)
        p = probe(host)
        out("  隧道 %-6s TLS %-7s HTTP %-5s %s" % (
            ("%dms" % p["tunnel_ms"]) if p["tunnel_ms"] is not None else "-",
            ("%dms" % p["tls_ms"]) if p["tls_ms"] is not None else "-",
            p["status"] if p["status"] else "-",
            ("server=%s" % p["server"]) if p["server"] else ""))
        out("  判定: " + verdict(p))


def cmd_site(args):
    rules = load_rules()
    main = args.domain.lstrip(".")
    base = registrable(main)
    url = args.url or ("https://%s/" % (main if main.count(".") > 1 else "www." + main))

    out("=" * 72)
    out("站点 %s" % main)
    out("=" * 72)

    out("\n[1] 主域名规则")
    main_rule = show_rule(main, rules)
    main_proxy = main_rule.get("proxy") if main_rule else None

    out("\n[2] 主域名连通性")
    p = probe(main)
    out("  隧道 %-6s TLS %-7s HTTP %-5s %s" % (
        ("%dms" % p["tunnel_ms"]) if p["tunnel_ms"] is not None else "-",
        ("%dms" % p["tls_ms"]) if p["tls_ms"] is not None else "-",
        p["status"] if p["status"] else "-",
        ("server=%s" % p["server"]) if p["server"] else ""))
    out("  判定: " + verdict(p))

    # ---- 域名发现：静态抓页面 + 可选实时观察
    hosts = {}
    out("\n[3] 页面引用的域名")
    try:
        html, status = fetch(url)
        for h, n in extract_hosts(html):
            hosts[h] = n
        out("  抓取 %s -> HTTP %s，抽到 %d 个域名" % (url, status, len(hosts)))
    except Exception as e:
        out("  抓取失败（%s）—— Cloudflare 质询或站点拒绝，"
            "改用 --watch 让浏览器去加载" % e.__class__.__name__)

    if args.watch:
        out("\n  实时观察 %d 秒，请现在用浏览器打开该站点..." % args.watch)
        seen = watch(args.watch)
        related = {h: v for h, v in seen.items()
                   if base in h or h in hosts}
        for h in related:
            hosts.setdefault(h, 0)
        out("  连接流里捕获 %d 个相关域名" % len(related))

    if not hosts:
        out("  没拿到域名，跳过后续分析")
        return

    # ---- 逐个查规则，挑出问题
    out("\n[4] 逐域名路由（按引用次数排序）")
    fallback, mismatch = [], []
    for h, n in sorted(hosts.items(), key=lambda kv: -kv[1])[:args.top]:
        hit, _, _ = match_domain(h, rules)
        if not hit:
            continue
        _, r = hit
        proxy = r.get("proxy")
        flag = ""
        if r.get("type") == "Match":
            fallback.append(h)
            flag = "  [落兜底]"
        elif main_proxy and proxy != main_proxy and (base in h or n >= 5):
            mismatch.append((h, proxy))
            flag = "  [出口与主站不同]"
        out("  %-42s x%-4d %-13s -> %s%s" %
            (h[:42], n, r.get("type"), proxy, flag))

    # ---- 结论
    out("\n[5] 结论")
    if not fallback and not mismatch:
        out("  没发现路由异常。若页面仍不正常，问题在浏览器层"
            "（广告拦截插件、缓存），不在 Clash。")
    if fallback:
        out("  落兜底的域名 %d 个，其中可能有该站的 CDN：" % len(fallback))
        for h in fallback[:12]:
            out("     " + h)
    if mismatch:
        out("  出口与主站不一致 %d 个 —— 主站能连而它们连不上时，"
            "页面会缺样式或缺图：" % len(mismatch))
        for h, pr in mismatch[:12]:
            out("     %-42s -> %s" % (h[:42], pr))

    if args.dns:
        out("\n[6] DNS 地理合理性")
        for h in [main] + [x for x, _ in sorted(hosts.items(), key=lambda kv: -kv[1])[:3]]:
            a, b = dns_via_mihomo(h), dns_via_alidns(h)
            same = set(a) & set(b)
            out("  %-38s mihomo=%-34s 阿里=%s%s" %
                (h[:38], ",".join(a[:2]) or "-", ",".join(b[:2]) or "-",
                 "" if same or not (a and b) else "   [差异大，可能解析到境外节点]"))


# ---------------------------------------------------------------- 常驻监控
# 设计前提：噪音是这类工具唯一的死因。
# 随便一个页面就拉三五十个域名，其中一大半是广告统计，落兜底完全正常。
# 全报出来一天几百条，三天就会被关掉。
#
# 所以只在【连接真的失败】时才出声 —— 成功的一律不吭气，不管它走哪条链。
# 失败时再把同期活跃的其他域名一并列出，出口不一致会自己浮出来。
# spankbang 那次就是这个形状：主站走自动选择能通，CDN 走住宅链连不上。
LOG_DIR_DEFAULT = os.path.join(os.environ.get("LOCALAPPDATA", "."), "netwatch")


def prune_logs(log_dir, keep_days):
    """按天滚动删除。这份日志等于一份比浏览器历史还细的浏览记录，
    明文躺在磁盘上，所以默认只留几天。"""
    if not os.path.isdir(log_dir):
        return
    cutoff = time.time() - keep_days * 86400
    for fn in os.listdir(log_dir):
        if not fn.startswith("netwatch-") or not fn.endswith(".jsonl"):
            continue
        p = os.path.join(log_dir, fn)
        try:
            if os.path.getmtime(p) < cutoff:
                os.remove(p)
        except OSError:
            pass


def cmd_daemon(args):
    log_dir = args.log or LOG_DIR_DEFAULT
    os.makedirs(log_dir, exist_ok=True)
    prune_logs(log_dir, args.keep_days)

    rules = load_rules()
    rules_at = time.time()
    active = {}        # conn id -> 记录
    recent = []        # (ts, host, chain) 滚动窗口，用于给失败提供上下文
    silenced = {}      # host -> 上次报告时间，同一域名不反复刷屏
    last_prune = time.time()

    out("netwatch 常驻监控已启动")
    out("  控制器 %s   代理端口 %d" % (CTRL, MIXED_PORT))
    out("  日志   %s（保留 %d 天）" % (log_dir, args.keep_days))
    out("  策略   只报连接失败；成功的一律静默")
    out("  Ctrl+C 退出\n")

    while True:
        now = time.time()

        # 配置可能被热重载，规则要跟着刷新
        if now - rules_at > 300:
            try:
                rules = load_rules()
                rules_at = now
            except Exception:
                pass

        if now - last_prune > 3600:
            prune_logs(log_dir, args.keep_days)
            last_prune = now

        try:
            conns = api("/connections", timeout=8).get("connections") or []
        except Exception:
            time.sleep(args.interval)
            continue

        cur = set()
        for c in conns:
            cid = c.get("id")
            if not cid:
                continue
            cur.add(cid)
            meta = c.get("metadata") or {}
            host = meta.get("host") or meta.get("destinationIP") or ""
            if not host:
                continue
            dl = c.get("download", 0) or 0
            if cid in active:
                active[cid]["download"] = dl
            else:
                chain = " <- ".join(c.get("chains") or [])
                active[cid] = {"host": host, "chain": chain, "download": dl,
                               "rule": c.get("rule"), "payload": c.get("rulePayload"),
                               "ip": meta.get("destinationIP", ""), "seen": now}
                recent.append((now, host, chain))

        recent = [r for r in recent if now - r[0] < 10]

        for cid in list(active):
            if cid in cur:
                continue
            rec = active.pop(cid)
            lived = now - rec["seen"]
            # 一个字节都没收到 + 活过一瞬间 = 握手失败或被拒，不是正常的短连接
            if rec["download"] > 0 or lived < 0.3:
                continue
            host = rec["host"]
            if now - silenced.get(host, 0) < args.silence:
                continue
            silenced[host] = now

            # 同期上下文里要滤掉裸 IP：那多半是代理链自己的入口/出口连接，
            # 对判断没帮助，只会把真正有用的域名挤下去
            others = {}
            for ts, h, ch in recent:
                if h == host or now - ts >= 6 or IP_RE.match(h):
                    continue
                others.setdefault(ch, set()).add(h)

            entry = {"ts": time.strftime("%Y-%m-%d %H:%M:%S"), "host": host,
                     "rule": rec["rule"], "payload": rec["payload"],
                     "chain": rec["chain"], "lived_ms": int(lived * 1000),
                     "concurrent": {k: sorted(v)[:6] for k, v in others.items()}}
            try:
                with io.open(os.path.join(
                        log_dir, "netwatch-%s.jsonl" % time.strftime("%Y%m%d")),
                        "a", encoding="utf-8") as f:
                    f.write(json.dumps(entry, ensure_ascii=False) + "\n")
            except OSError:
                pass

            if args.quiet:
                continue
            out("[%s] 连接失败  %s" % (entry["ts"], host))
            out("    规则 %s %s  ->  %s" %
                (rec["rule"], rec["payload"] or "", rec["chain"] or "?"))
            if rec["rule"] == "Match":
                out("    ↑ 落兜底，没有专门规则 —— 很可能是该站的 CDN 漏配了")
            # 同期其他出口：出口不一致时这里会直接显形
            for ch, hs in others.items():
                if ch and ch != rec["chain"]:
                    out("    同期走 %s 的域名: %s" % (ch, ", ".join(sorted(hs)[:4])))
            out("")

        time.sleep(args.interval)


def cmd_recent(args):
    """看最近记录下来的异常，不用一直盯着终端"""
    log_dir = args.log or LOG_DIR_DEFAULT
    if not os.path.isdir(log_dir):
        out("还没有日志：%s" % log_dir)
        return
    files = sorted(fn for fn in os.listdir(log_dir)
                   if fn.startswith("netwatch-") and fn.endswith(".jsonl"))
    rows = []
    for fn in files[-args.days:]:
        try:
            for line in io.open(os.path.join(log_dir, fn), encoding="utf-8"):
                rows.append(json.loads(line))
        except (OSError, ValueError):
            continue
    if not rows:
        out("最近 %d 天没有记录到连接失败" % args.days)
        return
    out("最近 %d 天共 %d 条失败记录\n" % (args.days, len(rows)))
    for e in rows[-args.limit:]:
        out("%s  %s" % (e["ts"], e["host"]))
        out("    %s %s -> %s" % (e.get("rule"), e.get("payload") or "",
                                 e.get("chain") or "?"))
        for ch, hs in (e.get("concurrent") or {}).items():
            if ch and ch != e.get("chain"):
                out("    同期走 %s: %s" % (ch, ", ".join(hs[:4])))
        out("")


def main():
    ap = argparse.ArgumentParser(description="Clash / mihomo 路由诊断")
    ap.add_argument("--ascii", action="store_true", help="输出转义，规避控制台编码问题")
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("site", help="分析一个站点牵扯的全部域名及各自出口")
    s.add_argument("domain")
    s.add_argument("--url", help="指定要抓取的页面，默认 https://<domain>/")
    s.add_argument("--watch", type=int, default=0, help="同时观察连接流 N 秒")
    s.add_argument("--top", type=int, default=25, help="最多列出多少个域名")
    s.add_argument("--dns", action="store_true", help="附带 DNS 地理合理性对比")
    s.set_defaults(func=cmd_site)

    c = sub.add_parser("check", help="分层探测若干域名")
    c.add_argument("domains", nargs="+")
    c.set_defaults(func=cmd_check)

    r = sub.add_parser("rules", help="只看规则命中")
    r.add_argument("domains", nargs="+")
    r.set_defaults(func=cmd_rules)

    d = sub.add_parser("daemon", help="常驻监控，只在连接失败时出声")
    d.add_argument("--interval", type=float, default=0.5, help="轮询间隔秒")
    d.add_argument("--log", help="日志目录，默认 %%LOCALAPPDATA%%\\netwatch")
    d.add_argument("--keep-days", type=int, default=3, help="日志保留天数")
    d.add_argument("--silence", type=int, default=600,
                   help="同一域名多少秒内不重复报告")
    d.add_argument("--quiet", action="store_true", help="只写日志不打印")
    d.set_defaults(func=cmd_daemon)

    n = sub.add_parser("recent", help="查看最近记录到的失败")
    n.add_argument("--log", help="日志目录")
    n.add_argument("--days", type=int, default=3)
    n.add_argument("--limit", type=int, default=30)
    n.set_defaults(func=cmd_recent)

    args = ap.parse_args()
    setup_stdout(args.ascii)

    global CTRL, SECRET, MIXED_PORT
    CTRL, SECRET, MIXED_PORT = read_runtime_config()
    args.func(args)


if __name__ == "__main__":
    main()

import html as htmllib
import os
import re
import subprocess
import sys
import time
from urllib.parse import urljoin

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi"])
from curl_cffi import requests as creq

MAIN = "https://www.axa.com.hk/zh/fulfilment-ratios-and-total-value-ratios"
PCT = re.compile(r"\d+(?:\.\d+)?\s*%")
out = []


def say(s=""):
    print(s)
    out.append(s)


def get(url):
    last = ""
    for i in range(3):
        try:
            r = creq.get(url, impersonate="chrome", timeout=60)
            if r.status_code == 200:
                return r.content.decode("utf-8", "replace"), "200"
            last = f"狀態 {r.status_code}"
        except Exception as e:
            last = str(e)[:80]
        time.sleep(2 * (i + 1))
    return "", last


def clean(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def cells(tr):
    row = []
    for attrs, inner in re.findall(r"<t[hd]\b([^>]*)>(.*?)</t[hd]>", tr, flags=re.S | re.I):
        m = re.search(r"colspan\s*=\s*[\"']?(\d+)", attrs, flags=re.I)
        row += [clean(inner)[:28]] * (int(m.group(1)) if m else 1)
    return row


def report(url, body, full):
    title = re.search(r"<title[^>]*>(.*?)</title>", body, flags=re.S | re.I)
    tables = list(re.finditer(r"<table\b[^>]*>(.*?)</table>", body, flags=re.S | re.I))
    say(f"標題：{clean(title.group(1))[:80] if title else '—'}")
    say(f"長度 {len(body)}｜表格 {len(tables)} 個｜百分比 {len(PCT.findall(clean(body)))} 個")
    heads = [clean(h)[:50] for h in re.findall(r"<h[1-4]\b[^>]*>(.*?)</h[1-4]>", body, flags=re.S | re.I)]
    say("小標題：" + "｜".join(h for h in heads if h)[:600])
    if not tables:
        hints = set(re.findall(r"[\"']([^\"']*\.(?:json|pdf|xlsx)[^\"']*)[\"']", body)[:15])
        hints |= set(re.findall(r"<iframe[^>]*src=[\"']([^\"']+)", body)[:5])
        hints |= set(re.findall(r"[\"'](/[^\"']*api/[^\"']*)[\"']", body)[:10])
        say("沒有表格，可能的資料來源：")
        for h in sorted(hints):
            say("    " + h[:200])
        return
    end = 0
    for ti, m in enumerate(tables):
        before = clean(body[max(end, m.start() - 3000):m.start()])[-120:]
        end = m.end()
        rows = [cells(tr) for tr in re.findall(r"<tr\b.*?</tr>", m.group(1), flags=re.S | re.I)]
        rows = [r for r in rows if r]
        width = max((len(r) for r in rows), default=0)
        say(f"  表格 {ti}：{len(rows)} 行 × {width} 欄｜前文：{before}")
        for r in rows[:(8 if full else 2)]:
            say("      " + " | ".join(r))
        if full and len(rows) > 8:
            say(f"      …（尚有 {len(rows) - 8} 行）")
            say("      " + " | ".join(rows[-1]))


body, status = get(MAIN)
say(f"=== 總頁 === {MAIN}｜{status}")
if not body:
    say("讀不到總頁")
else:
    report(MAIN, body, True)
    links = []
    for href in re.findall(r"href=[\"']([^\"'#]*fulfil[^\"'#]*)[\"']", body, flags=re.I):
        u = urljoin(MAIN, htmllib.unescape(href)).split("?")[0].rstrip("/")
        if "/zh/" in u and u != MAIN and u not in links:
            links.append(u)
    say()
    say(f"=== 產品頁連結 {len(links)} 個 ===")
    for u in links:
        say("  " + u)
    for i, u in enumerate(links[:120]):
        sub, status = get(u)
        say()
        say(f"=== 產品頁 {i + 1}/{len(links)} === {u}｜{status}")
        if sub:
            report(u, sub, i < 6)
        time.sleep(0.5)

os.makedirs("output", exist_ok=True)
with open("output/axa-detect.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out))

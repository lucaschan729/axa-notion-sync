import collections
import html as htmllib
import json
import os
import re
import subprocess
import sys
import time
from urllib.parse import quote

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi"])
from curl_cffi import requests as creq

BASE = "https://www.axa.com.hk/zh/"
MAIN = BASE + "fulfilment-ratios-and-total-value-ratios"
KNOWN = ["fulfilment-ratios-total-value-ratios-maxx-savings-series",
         "fulfilment-ratios-total-value-ratios-wealth-genius-income-plan"]
PRISMIC = "https://hk-axa-web-2024.cdn.prismic.io/api/v2"
SLUG = re.compile(r"fulfill?ment-ratios?-[a-z0-9-]+")
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
            if r.status_code in (403, 404):
                break
        except Exception as e:
            last = str(e)[:80]
        time.sleep(2 * (i + 1))
    return "", last


def clean(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def flat(s):
    return re.sub(r"\s+", " ", s)


def walk(obj, path, hits):
    if isinstance(obj, str):
        if PCT.search(obj) or "<table" in obj or "保單年度" in obj:
            hits.append((path, obj[:140]))
    elif isinstance(obj, dict):
        for k, v in obj.items():
            walk(v, f"{path}.{k}", hits)
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            walk(v, f"{path}[{i}]", hits)


def inspect(body, full):
    title = re.search(r"<title[^>]*>(.*?)</title>", body, flags=re.S | re.I)
    tables = re.findall(r"<table\b[^>]*>(.*?)</table>", body, flags=re.S | re.I)
    say(f"標題：{clean(title.group(1))[:80] if title else '—'}")
    say(f"長度 {len(body)}｜表格 {len(tables)} 個｜可見文字百分比 {len(PCT.findall(clean(body)))} 個"
        f"｜原始碼「保單年度」{body.count('保單年度')} 次｜「%」{body.count('%')} 次")
    nd = re.search(r"<script[^>]*id=[\"']__NEXT_DATA__[\"'][^>]*>(.*?)</script>", body, flags=re.S)
    say(f"__NEXT_DATA__：{'有，長度 ' + str(len(nd.group(1))) if nd else '沒有'}"
        f"｜__next_f：{body.count('self.__next_f.push')} 段")
    if nd:
        try:
            data = json.loads(nd.group(1))
            say(f"  page={data.get('page')}｜buildId={data.get('buildId')}")
            pp = data.get("props", {}).get("pageProps", {})
            say(f"  pageProps 鍵：{list(pp)[:20]}")
            hits = []
            walk(pp, "pageProps", hits)
            say(f"  含百分比／表格／保單年度的字串：{len(hits)} 個")
            for p, v in hits[:(30 if full else 5)]:
                say(f"    {p} = {flat(v)}")
        except Exception as e:
            say(f"  JSON 解析失敗：{str(e)[:80]}")
    if full:
        for ti, t in enumerate(tables[:3]):
            rows = re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I)
            say(f"  表格 {ti}：{len(rows)} 行")
            for tr in rows[:6]:
                cs = re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", tr, flags=re.S | re.I)
                say("      " + " | ".join(clean(c)[:26] for c in cs))
        for kw in ["保單年度", "報告年度之表現", "繕發年份", "生效"]:
            pos = [m.start() for m in re.finditer(re.escape(kw), body)]
            for p in pos[:2]:
                say(f"  「{kw}」附近原始碼：{flat(body[max(0, p - 200):p + 700])}")
    return sorted(set(SLUG.findall(body)))


body, status = get(MAIN)
say(f"=== 總頁 === {MAIN}｜{status}")
slugs = []
if body:
    slugs = [s for s in inspect(body, True) if s != "fulfilment-ratios-and-total-value-ratios"]
say()
say(f"=== 原始碼內的產品頁名稱 {len(slugs)} 個 ===")
for s in slugs:
    say("  " + s)

todo = list(dict.fromkeys(KNOWN + slugs))
for i, s in enumerate(todo[:60]):
    sub, status = get(BASE + s)
    say()
    say(f"=== 產品頁 {i + 1}/{len(todo)} === {BASE + s}｜{status}")
    if sub:
        inspect(sub, i < 3)
    time.sleep(0.5)

say()
say(f"=== Prismic === {PRISMIC}")
api, status = get(PRISMIC)
say(f"狀態：{status}")
if api:
    try:
        info = json.loads(api)
        ref = next(r["ref"] for r in info["refs"] if r.get("isMasterRef"))
        say(f"類型：{list(info.get('types', {}))[:60]}")
        say(f"語言：{[l.get('id') for l in info.get('languages', [])]}")
        found, types, total = [], collections.Counter(), 0
        for page in range(1, 31):
            txt, st = get(f"{PRISMIC}/documents/search?ref={quote(ref)}&lang=*&pageSize=100&page={page}")
            if not txt:
                say(f"第 {page} 頁讀取失敗：{st}")
                break
            res = json.loads(txt)
            total = res.get("total_results_size", 0)
            for d in res.get("results", []):
                types[d.get("type")] += 1
                uid = d.get("uid") or ""
                if "fulfil" in uid or "ratio" in uid:
                    found.append(d)
            if page >= res.get("total_pages", 1):
                break
        say(f"文件總數：{total}｜各類型：{dict(types)}")
        say(f"名稱含 fulfil／ratio 的文件：{len(found)} 個")
        for d in found:
            say(f"  {d.get('type')}｜{d.get('lang')}｜{d.get('uid')}")
        zh = [d for d in found if "zh" in (d.get("lang") or "")] or found
        for d in zh[:2]:
            say()
            say(f"--- 文件內容樣本：{d.get('uid')}｜{d.get('lang')} ---")
            say(f"data 鍵：{list(d.get('data', {}))}")
            hits = []
            walk(d.get("data", {}), "data", hits)
            say(f"含百分比／表格／保單年度的字串：{len(hits)} 個")
            for p, v in hits[:40]:
                say(f"    {p} = {flat(v)}")
            say(json.dumps(d.get("data", {}), ensure_ascii=False)[:3000])
    except Exception as e:
        say(f"Prismic 解析失敗：{str(e)[:200]}")

os.makedirs("output", exist_ok=True)
with open("output/axa-detect.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out))

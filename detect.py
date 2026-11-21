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
            last = str(e)[:60]
        time.sleep(2 * (i + 1))
    return "", last


def clean(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def tables_of(body):
    return list(re.finditer(r"<table\b[^>]*>(.*?)</table>", body, flags=re.S | re.I))


def next_data(body):
    m = re.search(r"<script[^>]*id=[\"']__NEXT_DATA__[\"'][^>]*>(.*?)</script>", body, flags=re.S)
    return m.group(1) if m else ""


def walk(obj, path, hits):
    if isinstance(obj, str):
        if PCT.search(obj) or "<table" in obj or "保單年度" in obj:
            hits.append((path, re.sub(r"\s+", " ", obj)[:50]))
    elif isinstance(obj, dict):
        for k, v in obj.items():
            walk(v, f"{path}.{k}", hits)
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            walk(v, f"{path}[{i}]", hits)


def short(path):
    return re.sub(r"\[\d+\]", "[]", path)[-70:]


body, status = get(MAIN)
say(f"=== 總頁 === {status}")
slugs = [s for s in sorted(set(SLUG.findall(body))) if s != "fulfilment-ratios-and-total-value-ratios"]
say(f"產品頁名稱 {len(slugs)} 個｜NEXT_DATA {'有' if next_data(body) else '沒有'}"
    f"｜next_f {body.count('self.__next_f.push')} 段")

todo = list(dict.fromkeys(KNOWN + slugs))[:90]
pages, failed = {}, []
for s in todo:
    sub, st = get(BASE + s)
    if sub:
        pages[s] = sub
    else:
        failed.append(f"{s[-30:]}({st})")
    time.sleep(0.3)
counts = {s: len(tables_of(b)) for s, b in pages.items()}
pcts = {s: len(PCT.findall(clean(b))) for s, b in pages.items()}
say()
say("=== 產品頁統計 ===")
say(f"讀到 {len(pages)}/{len(todo)} 頁｜有表格 {sum(1 for c in counts.values() if c)} 頁"
    f"｜表格共 {sum(counts.values())} 個｜可見百分比共 {sum(pcts.values())} 個")
say(f"每頁表格數分佈：{dict(sorted(collections.Counter(counts.values()).items()))}")
say(f"沒有表格的頁：{[s[-30:] for s, c in counts.items() if not c][:6]}")
say(f"讀取失敗：{failed[:6]}")

sample = next((s for s in todo if counts.get(s)), next(iter(pages), None))
if sample:
    b = pages[sample]
    say()
    say(f"=== 樣本頁 === …{sample[-40:]}")
    say(f"表格 {counts[sample]} 個｜百分比 {pcts[sample]} 個｜NEXT_DATA {'有' if next_data(b) else '沒有'}"
        f"｜next_f {b.count('self.__next_f.push')} 段｜原始碼「保單年度」{b.count('保單年度')} 次")
    end = 0
    for ti, m in enumerate(tables_of(b)[:3]):
        before = clean(b[max(end, m.start() - 2500):m.start()])[-45:]
        end = m.end()
        trs = re.findall(r"<tr\b.*?</tr>", m.group(1), flags=re.S | re.I)
        span = len(re.findall(r"(?:col|row)span", m.group(1), flags=re.I))
        say(f"表{ti}：{len(trs)} 行｜合併格 {span} 個｜前文：{before}")
        for tr in trs[:4]:
            cs = re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", tr, flags=re.S | re.I)
            say(f"  [{len(cs)}格] " + "|".join(clean(c)[:11] for c in cs[:9]))
    nd = next_data(b)
    if nd:
        try:
            hits = []
            walk(json.loads(nd).get("props", {}).get("pageProps", {}), "pp", hits)
            say(f"NEXT_DATA 內含數字的字串 {len(hits)} 個")
            for p, v in hits[:5]:
                say(f"  {short(p)} = {v}")
        except Exception as e:
            say(f"NEXT_DATA 解析失敗：{str(e)[:60]}")
    if not counts[sample]:
        p = b.find("保單年度")
        if p >= 0:
            say("「保單年度」附近：" + re.sub(r"\s+", " ", b[max(0, p - 150):p + 350]))

say()
api, status = get(PRISMIC)
say(f"=== Prismic === {status}")
if api:
    try:
        ref = next(r["ref"] for r in json.loads(api)["refs"] if r.get("isMasterRef"))
        found, total = [], 0
        for page in range(1, 31):
            txt, st = get(f"{PRISMIC}/documents/search?ref={quote(ref)}&lang=*&pageSize=100&page={page}")
            if not txt:
                say(f"第 {page} 頁讀取失敗：{st}")
                break
            res = json.loads(txt)
            total = res.get("total_results_size", 0)
            found += [d for d in res.get("results", []) if "fulfil" in (d.get("uid") or "")]
            if page >= res.get("total_pages", 1):
                break
        say(f"文件總數 {total}｜fulfil 文件 {len(found)} 個")
        say(f"類型：{dict(collections.Counter(d.get('type') for d in found))}")
        say(f"語言：{dict(collections.Counter(d.get('lang') for d in found))}")
        zh = [d for d in found if d.get("uid") in KNOWN and "cn" not in (d.get("lang") or "")] or found
        if zh:
            d = zh[0]
            data = d.get("data", {})
            say(f"樣本：{d.get('lang')}｜…{(d.get('uid') or '')[-30:]}")
            say(f"data 鍵：{list(data)[:12]}")
            slices = [x for v in data.values() if isinstance(v, list) for x in v
                      if isinstance(x, dict) and x.get("slice_type")]
            say(f"slice 類型：{dict(collections.Counter(x['slice_type'] for x in slices))}")
            hits = []
            walk(data, "data", hits)
            say(f"含數字的字串 {len(hits)} 個")
            for p, v in hits[:6]:
                say(f"  {short(p)} = {v}")
    except Exception as e:
        say(f"Prismic 解析失敗：{str(e)[:120]}")

os.makedirs("output", exist_ok=True)
with open("output/axa-detect.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out))

import collections
import datetime
import html as htmllib
import json
import math
import os
import re
import subprocess
import sys
import time

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi", "openpyxl", "requests"])
import requests
from curl_cffi import requests as creq
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties

BASE = "https://www.axa.com.hk/zh/"
MAIN_SLUG = "fulfilment-ratios-and-total-value-ratios"
TVR, FR = "總價值比率", "分紅實現率"
LABELS = [TVR, FR]
DEFS = {
    TVR: "總價值比率＝實際總價值 ÷ 銷售時建議書所示的總價值（包括保證及非保證利益）。100% 即與銷售時一致，高於 100% 即比銷售時多。定義以安盛網站的說明為準。",
    FR: "分紅實現率＝實際非保證利益 ÷ 銷售時建議書所示的非保證利益。100% 即與預期一致，高於 100% 即比預期多。定義以安盛網站的說明為準。",
}
LEGEND = "綠色＝100% 或以上；紅色＝低於 80%；「/」＝沒有數字（網站顯示「不適用」）。"
AVG_NOTE = "「持滿N年平均值」＝保單年度N及之後所有數字的平均；「當年平均值」＝該年度一欄所有數字的平均。平均值為固定數字，不會隨篩選變動。"
SLUG = re.compile(r"fulfill?ment-ratios?-[a-z0-9-]+")
PCT = re.compile(r"(\d+(?:\.\d+)?)\s*%")
YEARH = re.compile(r"第\s*(\d+)\s*個\s*保單年度")
TABLE = re.compile(r"<table\b[^>]*>(.*?)</table>", flags=re.S | re.I)
BLUE, GREY, LIGHT, DARK, PURPLE = "1F4E79", "7F7F7F", "BFBFBF", "262626", "7030A0"
PALETTE = ["4472C4", "548235", "C55A11", "008080", "806000", "A6195B"]
CAT_COLORS = {}
HEAD_FILL = PatternFill("solid", fgColor=BLUE)
INFO_FILL = PatternFill("solid", fgColor="F2F5FA")
SEP_FILL = PatternFill("solid", fgColor="DCE6F1")
GREEN_FILL = PatternFill("solid", fgColor="E2EFDA")
RED_FILL = PatternFill("solid", fgColor="FCE4E4")
GREEN_TEXT, RED_TEXT = "548235", "C00000"
THIN = Side(style="thin", color="D9D9D9")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
SEP_BORDER = Border(bottom=Side(style="thin", color="8EA9DB"))
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ROW_HEIGHT = 20
HEAD_ROW = 3
YEAR_W = 12
CHECK_KEY = "保單年度8"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def get(url):
    last = ""
    for i in range(3):
        try:
            r = creq.get(url, impersonate="chrome", timeout=60)
            if r.status_code == 200:
                return r.content.decode("utf-8", "replace")
            last = f"狀態 {r.status_code}"
        except Exception as e:
            last = str(e)[:80]
        time.sleep(3 * (i + 1))
    print(f"讀取失敗：{url}｜{last}")
    return ""


def clean(s):
    s = re.sub(r"<(script|style|sup)\b.*?</\1>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def span(attrs, name):
    m = re.search(name + r"\s*=\s*[\"']?(\d+)", attrs, flags=re.I)
    return max(1, int(m.group(1))) if m else 1


def grid(table):
    out, pend = [], {}
    for tr in re.findall(r"<tr\b.*?</tr>", table, flags=re.S | re.I):
        row = []

        def fill():
            while len(row) in pend:
                col = len(row)
                txt, left = pend[col]
                if left <= 1:
                    del pend[col]
                else:
                    pend[col] = (txt, left - 1)
                row.append(txt)

        for attrs, inner in re.findall(r"<t[hd]\b([^>]*)>(.*?)</t[hd]>", tr, flags=re.S | re.I):
            fill()
            txt, rs = clean(inner), span(attrs, "rowspan")
            for _ in range(span(attrs, "colspan")):
                if rs > 1:
                    pend[len(row)] = (txt, rs - 1)
                row.append(txt)
        fill()
        out.append(row)
    return out


def header_index(g):
    return next((i for i, row in enumerate(g) if sum(bool(YEARH.search(c)) for c in row) >= 2), None)


def year_key(h):
    m = YEARH.search(h)
    rest = h[m.end():]
    prior = any(w in h for w in ["或之前", "或以前", "及之前", "及以前"])
    plus = prior or "+" in re.split(r"[（(]", rest)[0]
    ys = re.findall(r"(?:19|20)\d{2}", rest)
    if not ys:
        sub = ""
    elif len(ys) == 1 and not plus:
        sub = ys[0]
    else:
        sub = f"{max(ys)}或之前"
    return f"保單年度{m.group(1)}{'+' if plus else ''}", sub


def product_info(body, slug, log):
    first = body.find("<table")
    first = first if first >= 0 else len(body)
    p = body.rfind("產品類別", 0, first)
    cat = ""
    if p >= 0:
        m = re.match(r"產品類別\s*[:：]?\s*(.+)", clean(body[p:first]))
        cat = m.group(1).strip()[:30] if m else ""
    heads = [clean(h) for h in re.findall(r"<h[1-6]\b[^>]*>(.*?)</h[1-6]>", body[:p if p >= 0 else first],
                                          flags=re.S | re.I)]
    heads = [h for h in heads if h]
    name = heads[-1] if heads else ""
    how = "小標題"
    if not name or "分紅實現率及總價值比率" in name:
        m = re.search(r"了解\s*(.+?)\s*的分紅實現率", body)
        name, how = (clean(m.group(1)), "頁面描述") if m else (slug.split("ratios-")[-1], "網址")
    log["name_from"][how] += 1
    return re.sub(r"\s*/\s*", " / ", name).strip(), cat


def classify(label, table_text, pre, log):
    for s in (label, table_text):
        a, b = "總價值" in s, ("分紅實現率" in s or "紅利" in s)
        if a != b:
            return TVR if a else FR
    log["ambiguous"] += 1
    return TVR if pre.rfind("總價值比率") >= pre.rfind("分紅實現率") else FR


def parse_page(slug, body, grids, shared, ysubs, log):
    name, cat = product_info(body, slug, log)
    rows, seen, used, end = [], set(), 0, 0
    for m, g in zip(TABLE.finditer(body), grids):
        start, end = max(end, m.start() - 3000), m.end()
        hi = header_index(g)
        if hi is None:
            continue
        if json.dumps(g, ensure_ascii=False) in shared:
            log["shared"] += 1
            continue
        idx = [i for i, c in enumerate(g[hi]) if YEARH.search(c)]
        keys = []
        for i in idx:
            key, sub = year_key(g[hi][i])
            if sub or key not in ysubs:
                ysubs[key] = sub or ysubs.get(key, "")
            keys.append(key)
        table_text = " ".join(" ".join(r) for r in g)
        pre, got = None, 0
        for row in g[hi + 1:]:
            if not row or not row[0]:
                continue
            vals = [row[i] if i < len(row) else "" for i in idx]
            if not any(PCT.search(c) or "不適用" in c for c in vals):
                continue
            if pre is None:
                pre = clean(body[start:m.start()])
            before = log["ambiguous"]
            label = classify(row[0], table_text, pre, log)
            if log["ambiguous"] > before:
                log["amb_rows"].append(f"{name[:24]}｜{row[0][:40]} → {label}")
            parts = re.split(r"\s+[–—\-－]\s+", row[0], maxsplit=1)
            cur, kind = (parts[0], parts[1]) if len(parts) == 2 else ("所有保單貨幣", parts[0])
            cur = "所有" if cur.startswith("所有") else (re.sub(r"\s*保單貨幣$", "", cur).strip() or cur)
            kind = "" if label == TVR else (re.sub(r"\s*(之|的)?\s*分紅實現率", "", kind).strip() or "分紅")
            nums = {}
            for k, c in zip(keys, vals):
                pm = PCT.search(c)
                nums[k] = float(pm.group(1)) / 100 if pm else "/"
            sig = (label, cur, kind, tuple(nums.items()))
            if sig in seen:
                continue
            seen.add(sig)
            rows.append({"label": label, "name": name, "cat": cat, "cur": cur, "kind": kind, "vals": nums})
            got += 1
        used += 1 if got else 0
    log["used"][used] += 1
    return rows


def pct_format(v):
    return "0%" if abs(v * 100 - round(v * 100)) < 1e-6 else "0.0%"


def numbers(rows):
    return [v for r in rows for v in r["vals"].values() if isinstance(v, float)]


def mean(nums):
    return sum(nums) / len(nums) if nums else None


def averages(rows, ycols):
    return [mean([r["vals"][k] for r in rows if isinstance(r["vals"].get(k), float)]) for k in ycols]


def cumulative(rows, ycols):
    return [mean([r["vals"][k] for r in rows for k in ycols[i:] if isinstance(r["vals"].get(k), float)])
            for i in range(len(ycols))]


def hold_label(i, ycols):
    if i == 0:
        return "所有年份平均值"
    n = ycols[i].replace("保單年度", "")
    if n.endswith("+"):
        return f"持滿{n[:-1]}年+平均值"
    if i + 1 < len(ycols) and ycols[i + 1].endswith("+"):
        return f"第{n}年平均值"
    return f"持滿{n}年平均值"


def text_width(s):
    return sum(2 if ord(ch) > 0x2E7F else 1 for ch in s)


def fit(texts, lo, hi, size=9):
    w = max((text_width(t) for t in texts), default=0) * size / 10 + 3
    return max(lo, min(hi, round(w, 1)))


def wrap_lines(s, width, size=9):
    cap = max(1.0, (width - 2) / (size / 10))
    return sum(max(1, math.ceil(text_width(p) / cap)) for p in str(s).split("\n"))


def head_cell(c, size=9):
    c.font = Font(bold=True, color="FFFFFF", size=size)
    c.fill = HEAD_FILL
    c.alignment = CENTER
    c.border = BOX


def banner(ws, text, ncol, size, height):
    c = ws.cell(row=1, column=1, value=text)
    c.font = Font(bold=True, size=size, color="FFFFFF")
    c.fill = HEAD_FILL
    c.alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    ws.row_dimensions[1].height = height


def info_row(ws, text, ncol, height):
    c = ws.cell(row=2, column=1, value=text)
    c.font = Font(size=8, color=GREY)
    c.fill = INFO_FILL
    c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol)
    ws.row_dimensions[2].height = height


def write_averages(ws, start, label, rows, ycols, ysubs, nlead, lead_w):
    last = nlead + len(ycols)
    kinds = list(dict.fromkeys(x["kind"] for x in rows))
    multi = len(kinds) > 1
    all_name = "所有分红" if multi else "所有產品"
    groups = [(all_name, rows)]
    if multi:
        groups += [(k, [x for x in rows if x["kind"] == k]) for k in kinds]
    groups = [(t, s) for t, s in groups if numbers(s)]

    def line(r, title, vals):
        c = ws.cell(row=r, column=nlead, value=title)
        c.font = Font(bold=True, size=9, color=DARK)
        c.alignment = CENTER
        c.border = BOX
        for j, v in enumerate(vals):
            c = ws.cell(row=r, column=nlead + 1 + j, value=v if v is not None else "/")
            c.font = Font(bold=True, size=9, color=DARK if v is not None else LIGHT)
            c.alignment = CENTER
            c.border = BOX
            if v is not None:
                c.number_format = "0.00%"
        ws.row_dimensions[r].height = max(22, 13 * wrap_lines(title, lead_w) + 7)

    def heading(r, title, heads):
        head_cell(ws.cell(row=r, column=nlead, value=title))
        for j, h in enumerate(heads):
            head_cell(ws.cell(row=r, column=nlead + 1 + j, value=h), 8)
        lines = max([wrap_lines(h, YEAR_W, 8) for h in heads] + [wrap_lines(title, lead_w)])
        ws.row_dimensions[r].height = 12 * lines + 10

    r = start
    c = ws.cell(row=r, column=nlead, value="平均值摘要")
    c.font = Font(bold=True, size=11, color=BLUE)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[r].height = 24
    r += 1
    heading(r, "分紅類型" if multi else label, [hold_label(i, ycols) for i in range(len(ycols))])
    for title, sub in groups:
        r += 1
        line(r, title, cumulative(sub, ycols))
    r += 1
    heading(r, "當年平均值", [ysubs.get(k, "") or k for k in ycols])
    r += 1
    line(r, all_name, averages(rows, ycols))
    r += 1
    c = ws.cell(row=r, column=nlead, value=AVG_NOTE)
    c.font = Font(size=8, color=GREY)
    c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.merge_cells(start_row=r, start_column=nlead, end_row=r, end_column=last)
    ws.row_dimensions[r].height = 12 * wrap_lines(AVG_NOTE, lead_w + YEAR_W * len(ycols), 8) + 8


def write_sheet(ws, label, rows, ycols, ysubs, year, today):
    kinds = list(dict.fromkeys(r["kind"] for r in rows))
    multi = len(kinds) > 1
    extras = [label, "當年平均值", "所有產品", "所有分红", "分紅類型", "平均值摘要 "]
    cols = [("产品名称", fit([r["name"] for r in rows] + ["产品名称"], 18, 46), "name"),
            ("产品种类", fit([r["cat"] for r in rows] + ["产品种类"], 12, 30), "cat"),
            ("货币", fit([r["cur"] for r in rows] + ["货币"] + ([] if multi else extras), 10, 24), "cur")]
    if multi:
        cols.append(("红利类型", fit(kinds + ["红利类型"] + extras, 12, 26), "kind"))
    heads = [c[0] for c in cols] + [(k.replace("保單年度", "保单年度") + "\n" + ysubs.get(k, "")).strip() for k in ycols]
    widths = [c[1] for c in cols] + [YEAR_W] * len(ycols)
    n = len(heads)
    info = "\n".join([DEFS[label], LEGEND, f"資料來源：AXA安盛香港網站（報告年度 {year}）｜更新日期：{today}"])
    banner(ws, f"安盛{label}（報告年度 {year}）", n, 12, 30)
    info_row(ws, info, n, 12 * wrap_lines(info, sum(widths), 8) + 10)
    for i, (h, w) in enumerate(zip(heads, widths), 1):
        head_cell(ws.cell(row=HEAD_ROW, column=i, value=h))
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[HEAD_ROW].height = 14 * max(wrap_lines(h, w) for h, w in zip(heads, widths)) + 10
    ri, prev = HEAD_ROW, None
    for r in rows:
        if r["name"] != prev:
            if prev is not None:
                ri += 1
                for ci in range(1, n + 1):
                    c = ws.cell(row=ri, column=ci)
                    c.fill = SEP_FILL
                    c.border = SEP_BORDER
                ws.row_dimensions[ri].height = 4
            prev = r["name"]
        ri += 1
        for ci, (_, w, key) in enumerate(cols, 1):
            c = ws.cell(row=ri, column=ci, value=r[key])
            c.border = BOX
            c.alignment = CENTER
            if key == "name":
                c.font = Font(bold=True, size=9, color=DARK)
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            elif key == "cat":
                c.font = Font(size=9, color=CAT_COLORS.get(r[key], DARK))
            elif key == "cur":
                c.font = Font(bold=True, size=9, color=PURPLE)
            else:
                c.font = Font(size=9, color=DARK)
        for j, k in enumerate(ycols):
            v = r["vals"].get(k, "/")
            c = ws.cell(row=ri, column=len(cols) + 1 + j, value=v)
            c.alignment = CENTER
            c.border = BOX
            if isinstance(v, float):
                c.number_format = pct_format(v)
                if v >= 0.9995:
                    c.fill = GREEN_FILL
                    c.font = Font(size=9, color=GREEN_TEXT)
                elif v < 0.7995:
                    c.fill = RED_FILL
                    c.font = Font(size=9, color=RED_TEXT)
                else:
                    c.font = Font(size=9, color=DARK)
            else:
                c.font = Font(size=9, color=LIGHT)
        lines = max(wrap_lines(r[key], w) for _, w, key in cols)
        ws.row_dimensions[ri].height = max(ROW_HEIGHT, 13 * lines + 7)
    ws.freeze_panes = f"B{HEAD_ROW + 1}"
    ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(n)}{max(ri, HEAD_ROW + 1)}"
    write_averages(ws, ri + 2, label, rows, ycols, ysubs, len(cols), cols[-1][1])
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = BLUE
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"


def write_overview(ws, sheets, year, today):
    ncol = 6
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "FFC000"
    first_w = fit([f"{year}{label}" for label, _ in sheets] + list(CAT_COLORS) + ["產品種類"], 22, 40, 11)
    for i, w in enumerate([first_w, 18, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def section(r, s):
        c = ws.cell(row=r, column=1, value=s)
        c.font = Font(bold=True, size=11, color=BLUE)
        c.alignment = Alignment(vertical="bottom")
        ws.row_dimensions[r].height = 22

    def body(r, i, v, fmt="0.00%", color=DARK):
        c = ws.cell(row=r, column=i, value=v if v is not None else "/")
        c.border = BOX
        c.font = Font(size=10, bold=i == 1, color=color)
        c.alignment = Alignment(horizontal="left" if i == 1 else "center", vertical="center",
                                indent=1 if i == 1 else 0)
        if isinstance(v, float):
            c.number_format = fmt
        ws.row_dimensions[r].height = 22

    banner(ws, f"安盛分紅表現概覽（報告年度 {year}）", ncol, 14, 34)
    info_row(ws, f"資料來源：AXA安盛香港網站｜更新日期：{today}", ncol, 20)
    ws.cell(row=2, column=1).font = Font(size=9, color=GREY)
    ws.row_dimensions[3].height = 10
    section(4, "資料概況")
    r = 5
    for i, h in enumerate(["工作表", "產品數", "行數", "有數字的格數", "整體平均值", "低於 80% 的比例"], 1):
        head_cell(ws.cell(row=r, column=i, value=h), 10)
    ws.row_dimensions[r].height = 22
    for label, rows in sheets:
        r += 1
        nums = numbers(rows)
        low = len([v for v in nums if v < 0.7995]) / len(nums) if nums else None
        name = f"{year}{label}"
        line = [name, len(set(x["name"] for x in rows)), len(rows), len(nums), mean(nums), low]
        for i, v in enumerate(line, 1):
            body(r, i, v, "0.0%" if i == 6 else "0.00%")
        link = ws.cell(row=r, column=1)
        link.hyperlink = Hyperlink(ref="", location=f"'{name}'!A1", display=name)
        link.value = name
        link.font = Font(size=10, bold=True, color=BLUE, underline="single")
    r += 1
    ws.row_dimensions[r].height = 10
    r += 1
    if CAT_COLORS:
        section(r, "按產品種類的整體平均值")
        r += 1
        for i, h in enumerate(["產品種類"] + [label for label, _ in sheets], 1):
            head_cell(ws.cell(row=r, column=i, value=h), 10)
        ws.row_dimensions[r].height = 22
        for cat, color in CAT_COLORS.items():
            r += 1
            line = [cat] + [mean(numbers([x for x in rows if x["cat"] == cat])) for _, rows in sheets]
            for i, v in enumerate(line, 1):
                body(r, i, v, color=color if i == 1 else DARK)
        r += 1
        ws.row_dimensions[r].height = 10
        r += 1
    section(r, "顏色說明")
    samples = [(1.0, GREEN_FILL, GREEN_TEXT, "100% 或以上：與銷售時的說明一致或更好"),
               (0.79, RED_FILL, RED_TEXT, "低於 80%：明顯低於銷售時的說明"),
               ("/", None, LIGHT, "沒有數字：網站顯示「不適用」")]
    for v, fill, color, desc in samples:
        r += 1
        c = ws.cell(row=r, column=1, value=v)
        c.font = Font(size=10, color=color)
        c.alignment = CENTER
        c.border = BOX
        if fill:
            c.fill = fill
        if isinstance(v, float):
            c.number_format = "0%"
        d = ws.cell(row=r, column=2, value=desc)
        d.font = Font(size=10, color=DARK)
        d.alignment = Alignment(vertical="center", indent=1)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncol)
        ws.row_dimensions[r].height = 22
    r += 1
    ws.row_dimensions[r].height = 10
    r += 1
    section(r, "使用提示")
    tips = [
        "・點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
        f"・每個工作表第 {HEAD_ROW} 行有篩選箭嘴，可按產品種類、貨幣或紅利類型篩選。",
        "・向右或向下捲動時，標題列和產品名稱會固定不動。",
        "・平均值摘要在每個工作表最底部。",
        "・手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。",
    ]
    for s in tips:
        r += 1
        c = ws.cell(row=r, column=1, value=s)
        c.font = Font(size=9, color=GREY)
        c.alignment = Alignment(vertical="center", indent=1)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        ws.row_dimensions[r].height = 20


def notion_upload(path, year):
    token = os.environ.get("NOTION_TOKEN", "").strip()
    raw = os.environ.get("NOTION_DATABASE_ID", "").strip()
    if not token or not raw:
        print("未設定 NOTION_TOKEN 或 NOTION_DATABASE_ID，略過 Notion 上傳")
        return True
    m = re.search(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", raw)
    ids = [m.group(0)] if m else re.findall(r"[0-9a-fA-F]{32}", raw.split("?")[0])
    if not ids:
        print("Notion 上傳失敗：NOTION_DATABASE_ID 內找不到資料庫 ID")
        return False
    dbid = ids[-1]
    head = {"Authorization": f"Bearer {token}", "Notion-Version": "2025-09-03"}

    def api(step, method, path_, **kw):
        r = requests.request(method, "https://api.notion.com/v1/" + path_, headers=head, timeout=120, **kw)
        if r.status_code >= 300:
            raise RuntimeError(f"{step}：狀態 {r.status_code}｜{r.text[:200]}")
        return r.json()

    try:
        try:
            ds_id = api("讀取資料庫", "GET", f"databases/{dbid}")["data_sources"][0]["id"]
        except Exception:
            ds_id = dbid
        props = api("讀取資料來源", "GET", f"data_sources/{ds_id}")["properties"]
        title_prop = next(n for n, p in props.items() if p["type"] == "title")
        file_prop = next((n for n, p in props.items() if p["type"] == "files"), None)

        def page_title(p):
            return "".join(t.get("plain_text", "") for t in p["properties"][title_prop]["title"]).strip()

        found = api("尋找年度頁面", "POST", f"data_sources/{ds_id}/query",
                    json={"filter": {"property": title_prop, "title": {"starts_with": year}}})["results"]
        page = (next((p for p in found if page_title(p) == f"{year}年"), None)
                or next((p for p in found if page_title(p) == year), None)
                or (found[0] if found else None))
        if page:
            print(f"Notion：找到「{page_title(page)}」頁面")
        else:
            page = api("建立年度頁面", "POST", "pages",
                       json={"parent": {"type": "data_source_id", "data_source_id": ds_id},
                             "properties": {title_prop: {"title": [{"text": {"content": f"{year}年"}}]}}})
            print(f"Notion：找不到年度頁面，已建立「{year}年」")
        fname = os.path.basename(path)
        ascii_name = f"AXA_{year}.xlsx"

        def is_auto(n):
            return n == fname or n.startswith("AXA_")

        up = api("建立上傳", "POST", "file_uploads",
                 json={"mode": "single_part", "filename": fname, "content_type": XLSX})
        with open(path, "rb") as f:
            api("傳送檔案", "POST", f"file_uploads/{up['id']}/send", files={"file": (ascii_name, f, XLSX)})
        new = {"name": fname, "type": "file_upload", "file_upload": {"id": up["id"]}}
        if file_prop:
            keep = []
            for x in page.get("properties", {}).get(file_prop, {}).get("files", []):
                if is_auto(x.get("name", "")):
                    continue
                if x.get("type") == "file":
                    keep.append({"name": x["name"], "type": "file", "file": {"url": x["file"]["url"]}})
                elif x.get("type") == "external":
                    keep.append({"name": x["name"], "type": "external", "external": {"url": x["external"]["url"]}})
            done = api("更新檔案欄", "PATCH", f"pages/{page['id']}",
                       json={"properties": {file_prop: {"files": keep + [new]}}})
            count = len(done.get("properties", {}).get(file_prop, {}).get("files", []))
            print(f"Notion：已放入「{file_prop}」欄｜保留原有檔案 {len(keep)} 個｜現有檔案 {count} 個")
            if count < len(keep) + 1:
                print("注意：檔案數目比預期少，請到 Notion 檢查原有檔案")
        else:
            kids = api("讀取頁面內容", "GET", f"blocks/{page['id']}/children?page_size=100")["results"]
            removed = 0
            for b in kids:
                if b.get("type") == "file" and is_auto(b["file"].get("name", "")):
                    api("刪除舊檔", "DELETE", f"blocks/{b['id']}")
                    removed += 1
            api("加入檔案", "PATCH", f"blocks/{page['id']}/children",
                json={"children": [{"type": "file", "file": new}]})
            print(f"Notion：已把檔案放在頁面內容最底｜替換舊自動檔 {removed} 個｜其他內容不變")
        return True
    except Exception as e:
        print(f"Notion 上傳失敗｜{str(e)[:300]}")
        return False


def main():
    body = get(BASE + MAIN_SLUG)
    if not body:
        print("讀不到總頁")
        sys.exit(1)
    slugs = [s for s in dict.fromkeys(SLUG.findall(body)) if s != MAIN_SLUG]
    ym = re.search(r"(\d{4})\s*報告年度", clean(body))
    year = ym.group(1) if ym else str(datetime.date.today().year)
    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d")
    pages, missing = {}, []
    for s in slugs:
        sub = get(BASE + s)
        if sub:
            pages[s] = (sub, [grid(m.group(1)) for m in TABLE.finditer(sub)])
        else:
            missing.append(s)
        time.sleep(0.3)
    count = collections.Counter()
    for sub, grids in pages.values():
        for sig in set(json.dumps(g, ensure_ascii=False) for g in grids if header_index(g) is not None):
            count[sig] += 1
    data_pages = sum(1 for _, grids in pages.values() if grids)
    shared = {sig for sig, c in count.items() if data_pages >= 4 and c > data_pages / 2}
    log = {"shared": 0, "ambiguous": 0, "amb_rows": [], "used": collections.Counter(),
           "name_from": collections.Counter()}
    ysubs, rows, empty = {}, [], []
    for s, (sub, grids) in pages.items():
        if not grids:
            continue
        got = parse_page(s, sub, grids, shared, ysubs, log)
        rows += got
        if not got:
            empty.append(s[-34:])
    ycols = sorted(ysubs, key=lambda k: (int(re.search(r"\d+", k).group()), k.endswith("+")))
    sheets = [(label, [r for r in rows if r["label"] == label]) for label in LABELS]
    cats = list(dict.fromkeys(r["cat"] for r in rows if r["cat"]))
    CAT_COLORS.update({cat: PALETTE[i % len(PALETTE)] for i, cat in enumerate(cats)})

    def show(vals):
        return " ".join(f"{v:.2%}" if v is not None else "/" for v in vals)

    print("=== 總結 ===")
    print(f"報告年度 {year}｜頁面名稱 {len(slugs)} 個｜讀到 {len(pages)} 頁｜有表格 {data_pages} 頁｜讀取失敗 {len(missing)} 頁")
    print(f"每頁採用的數據表數目：{dict(sorted(log['used'].items()))}｜略過共用示例表 {log['shared']} 個"
          f"｜行名無法判斷類別 {log['ambiguous']} 行")
    for s in log["amb_rows"][:10]:
        print(f"    無法判斷的行：{s}")
    print(f"產品名稱來源：{dict(log['name_from'])}｜沒有數據的頁：{empty[:6]}")
    print(f"保單年度欄 {len(ycols)} 個：", "、".join(f"{k}（{ysubs.get(k, '')}）" for k in ycols))
    for label, rs in sheets:
        cells = " ".join(str(sum(1 for r in rs if isinstance(r["vals"].get(k), float))) for k in ycols)
        print(f"{label}：產品 {len(set(r['name'] for r in rs))} 個｜行數 {len(rs)}｜有數字格數 {len(numbers(rs))}")
        print(f"    貨幣 {dict(collections.Counter(r['cur'] for r in rs))}")
        print(f"    類型 {dict(collections.Counter(r['kind'] or '—' for r in rs))}")
        print(f"    各年度數字格數：{cells}")
        print(f"    當年平均值：{show(averages(rs, ycols))}")
        print(f"    持滿平均值：{show(cumulative(rs, ycols))}")
    check = [r for r in rows if r["label"] == FR and isinstance(r["vals"].get(CHECK_KEY), float)]
    print(f"=== 分紅實現率 {CHECK_KEY}（{ysubs.get(CHECK_KEY, '')}）有數字的行：{len(check)} 行 ===")
    for r in check:
        print(f"  {r['vals'][CHECK_KEY]:.0%}｜{r['cur']}｜{r['kind']}｜{r['name'][:30]}")

    if not rows:
        print("完全讀不到資料，沒有產生 Excel")
        sys.exit(1)
    wb = Workbook()
    write_overview(wb.active, sheets, year, today)
    wb.active.title = "概覽"
    for label, rs in sheets:
        if rs:
            write_sheet(wb.create_sheet(f"{year}{label}"), label, rs, ycols, ysubs, year, today)
    os.makedirs("output", exist_ok=True)
    path = f"output/安盛總價值分紅實現率（{year}年）自動生成.xlsx"
    wb.save(path)
    print(f"已產生 {path}")
    if any(not rs for _, rs in sheets) or missing:
        print("其中一類沒有資料，或有頁面讀取失敗，不上傳 Notion")
        sys.exit(1)
    if not notion_upload(path, year):
        sys.exit(1)


main()

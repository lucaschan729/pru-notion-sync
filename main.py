import collections
import datetime
import difflib
import html as htmllib
import io
import os
import re
import zipfile

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

TOKEN = os.environ.get("NOTION_TOKEN", "").strip()
TARGET = os.environ.get("NOTION_DATABASE_ID", "").strip()
if not TOKEN or not TARGET:
    raise SystemExit("未設定 NOTION_TOKEN 或 NOTION_DATABASE_ID")
API = "https://api.notion.com/v1"
VERSION = "2025-09-03"
HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Notion-Version": VERSION,
    "Content-Type": "application/json",
}
URL_CASH = "https://www.prudential.com.hk/tc/general/performance/total-cash-value-ratio/"
URL_FULFIL = "https://www.prudential.com.hk/tc/general/performance/fulfillment-ratio/"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
AUTO = "自動生成"
HEAD_ROW = 3
FIRST = 4
LOW_LIMIT = 0.8
NAVY_HEX = "1F4E79"
PREFERRED = ["特別紅利現金價值", "歸原紅利現金價值", "累積非保證每月入息", "終期紅利現金價值"]


def solid(argb):
    return PatternFill(fill_type="solid", start_color=argb, end_color=argb)


DARK = solid("FF" + NAVY_HEX)
LIGHT = solid("FFDDEBF7")
PALE = solid("FFF4F7FB")
WHITE = solid("FFFFFFFF")
WHITE_BOLD = Font(bold=True, color="FFFFFFFF")
BLACK_BOLD = Font(bold=True, color="FF000000")
NAME_FONT = Font(bold=True, color="FF" + NAVY_HEX)
GREY_FONT = Font(color="FFB5B5B5")
LOW_FONT = Font(color="FFC00000")
TITLE_FONT = Font(bold=True, size=16, color="FFFFFFFF")
NOTE_FONT = Font(size=10, color="FF44546A")
SECTION_FONT = Font(bold=True, size=12, color="FF" + NAVY_HEX)
LINE = Side(style="thin", color="FFD0D7E2")
GROUP = Side(style="thin", color="FF7F9BBF")
BORDER = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)


# ---------- 讀取保誠網頁 ----------
def text(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", "", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def squash(s):
    return re.sub(r"\s+", "", s or "")


def attr(tag, name):
    m = re.search(r"\s" + name + r'="([^"]*)"', tag)
    return m.group(1) if m else ""


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def pct(s):
    m = re.match(r"^\s*(-?\d[\d,]*(?:\.\d+)?)\s*[%％]", s or "")
    return round(float(m.group(1).replace(",", "")) / 100, 4) if m else "/"


def year_label(cell):
    m = re.match(r"^\s*\d+\s*\+?\s*[\(（]\s*(\d{4})\s*([^\)）]*)[\)）]", cell)
    return m.group(1) + squash(m.group(2)) if m else ""


def bonus_type(label):
    return re.sub(r"(分紅實現率|實現率)$", "", squash(label))


def find_variant(chunk, tab_name):
    """找出「分期繳費／整付」等版本；回傳 None 代表這個表格沿用上一個表格的名稱。"""
    i = chunk.rfind("產品類別")
    if i < 0:
        return None
    pre, tn = squash(chunk[:i]), squash(tab_name)
    j = pre.rfind(tn) if tn else -1
    if j >= 0:
        tail = pre[j + len(tn):]
    else:
        m = re.search(r"[-–—－](分期繳費|整付)$", pre)
        tail = m.group(1) if m else ""
    tail = tail.strip("-–—－:：")
    return tail if len(tail) <= 12 else ""


def parse_table(t):
    grid = [[text(b) for b in re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", tr, flags=re.S | re.I)]
            for tr in re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I)]
    labels, start = [], 0
    for i, cells in enumerate(grid):
        ys = [y for y in (year_label(c) for c in cells) if y]
        if len(ys) >= 3:
            labels, start = ys, i + 1
            break
    rows, bad = [], 0
    for cells in grid[start:]:
        if not cells:
            continue
        if labels and len(cells) == len(labels) + 1:
            rows.append((cells[0], cells[1:]))
        else:
            bad += 1
    return labels, rows, bad


def load_page(url, what):
    r = requests.get(url, headers={"User-Agent": UA, "Accept-Language": "zh-HK,zh;q=0.9"}, timeout=60)
    if r.status_code != 200:
        raise SystemExit(f"讀取{what}頁面失敗：狀態 {r.status_code}")
    r.encoding = "utf-8"
    page = r.text
    tabs = {}
    for m in re.finditer(r'(<li\b[^>]*role="tab"[^>]*>)(.*?)</li>', page, flags=re.S):
        tabs[re.sub(r"-tab$", "", attr(m.group(1), "id"))] = text(m.group(2))
    panels = [(m.start(), re.sub(r"-tabpanel$", "", attr(m.group(0), "id")))
              for m in re.finditer(r'<div\b[^>]*role="tabpanel"[^>]*>', page)]
    recs, prev_end, last_pid, variant, bad_total = [], 0, None, "", 0
    for m in re.finditer(r"<table\b.*?</table>", page, flags=re.S | re.I):
        pos = m.start()
        pid, pstart, cat = "", 0, ""
        for start, i in panels:
            if start > pos:
                break
            pid, pstart = i, start
            if i.startswith("tabsextension"):
                cat = tabs.get(i, "")
        chunk = text(page[max(prev_end, pstart):pos])
        prev_end = m.end()
        if pid != last_pid:
            last_pid, variant = pid, ""
        tab_name = tabs.get(pid) or pid or "未知產品"
        v = find_variant(chunk, tab_name)
        if v is not None:
            variant = v
        caps = list(re.finditer(r"(\d{4})\s*報告年度的", chunk))
        year = caps[-1].group(1) if caps else ""
        label = chunk[caps[-1].end():].strip()[:40] if caps else ""
        labels, rows, bad = parse_table(m.group(0))
        bad_total += bad
        recs.append({"cat": cat, "name": tab_name + (f" - {variant}" if variant else ""),
                     "label": label, "year": year, "labels": labels, "rows": rows})
    print(f"已讀取{what}：分頁 {len(tabs)} 個｜表格 {len(recs)} 個｜格式不符而略過的行 {bad_total}")
    return recs


def report_year(recs):
    c = collections.Counter(rec["year"] for rec in recs if rec["year"])
    return c.most_common(1)[0][0] if c else ""


def build_rows(recs, cols, with_type, keep=None):
    rows, skipped = [], []
    for rec in recs:
        if keep and not keep(rec["label"]):
            skipped.append(f"{rec['name']}｜{rec['label']}")
            continue
        for cur, raws in rec["rows"]:
            m = dict(zip(rec["labels"], raws))
            fixed = [rec["name"], rec["cat"], squash(cur)]
            if with_type:
                fixed.append(bonus_type(rec["label"]) or "未分類")
            rows.append(fixed + [pct(m.get(c, "")) for c in cols])
    return rows, skipped


def last_is_before(cols):
    return len(cols) >= 2 and not re.fullmatch(r"\d{4}", cols[-1])


def make_header(cols, fixed):
    n = len(cols)
    out = list(fixed)
    for i, c in enumerate(cols):
        if i == n - 1 and last_is_before(cols):
            out.append(f"保单年度{n - 1}+\n{cols[-2]}之前")
        else:
            out.append(f"保单年度{i + 1}\n{c}")
    return out


# ---------- 平均值 ----------
def mean_or_slash(nums):
    return sum(nums) / len(nums) if nums else "/"


def build_summary(n_fixed, cols, groups):
    n = len(cols)
    width = n_fixed + n
    head = []
    for i in range(n):
        if i == 0:
            head.append("所有年份平均值")
        elif i == n - 1:
            head.append(f"持滿{n - 1}年+平均值")
        elif i == n - 2:
            head.append(f"第{n - 1}年平均值")
        else:
            head.append(f"持滿{i + 1}年平均值")
    cum = []
    for label, sub in groups:
        if not sub:
            continue
        line = [label]
        for start in range(n_fixed, width):
            line.append(mean_or_slash([r[c] for r in sub for c in range(start, width) if is_num(r[c])]))
        cum.append(line)
    year_head = list(cols)
    if last_is_before(cols):
        year_head[-1] = f"{cols[-2]}之前"
    all_label, everything = groups[0]
    year_line = [all_label] + [mean_or_slash([r[c] for r in everything if is_num(r[c])])
                               for c in range(n_fixed, width)]
    return {"head": head, "cum": cum, "year_head": year_head, "year": [year_line]}


def fmt_pct(v):
    return f"{v * 100:.2f}" if is_num(v) else "/"


# ---------- 生成 Excel ----------
def write_row(ws, r, col, values):
    for i, v in enumerate(values):
        ws.cell(row=r, column=col + i, value=v)


def dark(cell):
    cell.fill = DARK
    cell.font = WHITE_BOLD


def light(cell):
    cell.fill = LIGHT
    cell.font = BLACK_BOLD


def add_sheet(wb, title, heading, notes, header, rows, summary, head_label, n_fixed, widths):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = NAVY_HEX
    ws.sheet_view.showGridLines = False
    last_col = len(header)
    data_end = HEAD_ROW + len(rows)

    ws.cell(row=1, column=1, value=heading)
    ws.cell(row=2, column=1, value=notes)
    for rr in (1, 2):
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
        ws.cell(row=rr, column=1).alignment = LEFT
    ws.cell(row=1, column=1).font = TITLE_FONT
    ws.cell(row=1, column=1).fill = DARK
    ws.cell(row=2, column=1).font = NOTE_FONT
    ws.cell(row=2, column=1).fill = PALE
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 66

    write_row(ws, HEAD_ROW, 1, header)
    for c in range(1, last_col + 1):
        cell = ws.cell(row=HEAD_ROW, column=c)
        dark(cell)
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[HEAD_ROW].height = 42

    band = False
    for i, row in enumerate(rows):
        rr = FIRST + i
        first = i == 0 or rows[i - 1][0] != row[0]
        last = i == len(rows) - 1 or rows[i + 1][0] != row[0]
        if first and i > 0:
            band = not band
        base = PALE if band else WHITE
        border = Border(left=LINE, right=LINE,
                        top=GROUP if first else LINE,
                        bottom=GROUP if last else LINE)
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=rr, column=c, value=v)
            cell.border = border
            cell.alignment = LEFT if c == 1 else CENTER
            cell.fill = base
            if c == 1:
                cell.font = NAME_FONT
            elif c > n_fixed:
                if is_num(v):
                    cell.number_format = "0%"
                    if v < LOW_LIMIT:
                        cell.font = LOW_FONT
                else:
                    cell.font = GREY_FONT

    if rows:
        ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(last_col)}{data_end}"

    lc = n_fixed
    r = data_end + 2
    ws.cell(row=r, column=lc, value="平均值摘要").font = SECTION_FONT
    ws.row_dimensions[r].height = 26
    r += 1
    table_start = r
    write_row(ws, r, lc, [head_label] + summary["head"])
    r += 1
    value_rows = []
    for line in summary["cum"]:
        write_row(ws, r, lc, line)
        value_rows.append(r)
        r += 1
    year_head_row = r
    write_row(ws, r, lc, ["當年平均值"] + summary["year_head"])
    r += 1
    for line in summary["year"]:
        write_row(ws, r, lc, line)
        value_rows.append(r)
        r += 1
    table_end = r - 1
    for rr in range(table_start, table_end + 1):
        for c in range(lc, last_col + 1):
            cell = ws.cell(row=rr, column=c)
            cell.alignment = CENTER
            cell.border = BORDER
    for c in range(lc, last_col + 1):
        dark(ws.cell(row=table_start, column=c))
    dark(ws.cell(row=year_head_row, column=lc))
    for c in range(lc + 1, last_col + 1):
        light(ws.cell(row=year_head_row, column=c))
    for rr in (table_start, year_head_row):
        ws.row_dimensions[rr].height = 32
    for vr in value_rows:
        light(ws.cell(row=vr, column=lc))
        for c in range(lc + 1, last_col + 1):
            cell = ws.cell(row=vr, column=c)
            cell.number_format = "0.00%"
            cell.fill = WHITE
    note_row = table_end + 1
    ws.cell(row=note_row, column=lc,
            value="「持滿N年平均值」＝保單年度N及之後所有數字的平均；"
                  "「當年平均值」＝該年度一欄所有數字的平均。平均值為固定數字，不會隨篩選變動。")
    ws.merge_cells(start_row=note_row, start_column=lc, end_row=note_row, end_column=last_col)
    ws.cell(row=note_row, column=lc).font = NOTE_FONT
    ws.cell(row=note_row, column=lc).alignment = LEFT
    ws.row_dimensions[note_row].height = 30

    for i in range(1, last_col + 1):
        w = widths[i - 1] if i <= n_fixed else 13
        if i == lc:
            w = max(w, 16)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=FIRST, column=2)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"


# ---------- Notion ----------
def check(r, what):
    if r.status_code >= 300:
        print(f"{what} 失敗：{r.status_code} {r.text[:400]}")
        raise SystemExit(1)
    return r.json()


def resolve_source(raw):
    m = re.search(r"[0-9a-fA-F]{32}", raw.replace("-", ""))
    ident = m.group(0) if m else raw

    def first_source(db_id):
        r = requests.get(f"{API}/databases/{db_id}", headers=HEADERS, timeout=30)
        if r.status_code == 200:
            ds = r.json().get("data_sources") or []
            return ds[0]["id"] if ds else None
        return None

    found = first_source(ident)
    if found:
        return found
    r = requests.get(f"{API}/data_sources/{ident}", headers=HEADERS, timeout=30)
    if r.status_code == 200:
        return ident
    status = r.status_code
    r = requests.get(f"{API}/blocks/{ident}/children?page_size=100", headers=HEADERS, timeout=30)
    if r.status_code == 200:
        for b in r.json().get("results", []):
            if b.get("type") == "child_database":
                found = first_source(b["id"])
                if found:
                    return found
    print(f"找不到保誠的年份資料庫（狀態 {status}）。401＝金鑰不正確；"
          "404＝頁面未連接「AIA 自動更新」或 NOTION_DATABASE_ID 不正確。")
    raise SystemExit(1)


def query_pages(source):
    r = requests.post(f"{API}/data_sources/{source}/query", headers=HEADERS, json={"page_size": 100}, timeout=30)
    return check(r, "讀取年份資料庫").get("results", [])


def page_title(page):
    for prop in page.get("properties", {}).values():
        if prop.get("type") == "title":
            return "".join(t.get("plain_text", "") for t in prop.get("title", []))
    return ""


def file_blocks(page_id):
    r = requests.get(f"{API}/blocks/{page_id}/children?page_size=100", headers=HEADERS, timeout=30)
    out = []
    for block in check(r, "讀取頁面內容").get("results", []):
        if block["type"] == "file":
            info = block["file"]
            url = (info.get("file") or info.get("external") or {}).get("url")
            out.append((block["id"], info.get("name", ""), url))
    return out


def fix_xlsx(content):
    src = zipfile.ZipFile(io.BytesIO(content))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "xl/styles.xml":
                data = re.sub(r"<(\w+:)?family\b[^>]*/>", "", data.decode("utf-8")).encode("utf-8")
            dst.writestr(item.filename, data)
    return out.getvalue()


def load_manual_file(page_id):
    for _, name, url in file_blocks(page_id):
        if name.lower().endswith(".xlsx") and AUTO not in name and url:
            data = requests.get(url, timeout=60)
            data.raise_for_status()
            print("核對對象：", name)
            try:
                return load_workbook(io.BytesIO(data.content), data_only=True)
            except ValueError:
                return load_workbook(io.BytesIO(fix_xlsx(data.content)), data_only=True)
    return None


def upload_file(page_id, filename, content):
    old_auto = [bid for bid, name, _ in file_blocks(page_id) if AUTO in name]
    r = requests.post(f"{API}/file_uploads", headers=HEADERS,
                      json={"mode": "single_part", "filename": filename, "content_type": XLSX}, timeout=30)
    upload_id = check(r, "建立上傳")["id"]
    r = requests.post(f"{API}/file_uploads/{upload_id}/send",
                      headers={"Authorization": f"Bearer {TOKEN}", "Notion-Version": VERSION},
                      files={"file": (filename, content, XLSX)}, timeout=120)
    check(r, "傳送檔案")
    block = {"object": "block", "type": "file",
             "file": {"type": "file_upload", "file_upload": {"id": upload_id}}}
    r = requests.patch(f"{API}/blocks/{page_id}/children", headers=HEADERS, json={"children": [block]}, timeout=60)
    check(r, "加入檔案到頁面")
    for bid in old_auto:
        r = requests.delete(f"{API}/blocks/{bid}", headers=HEADERS, timeout=30)
        check(r, "刪除上一次的自動生成檔案")
    print(f"已上傳：{filename}（刪除舊的自動生成檔案 {len(old_auto)} 個）")


# ---------- 核對 ----------
def norm_val(v):
    if v is None or str(v).strip() == "":
        return "/"
    if is_num(v):
        return f"{v * 100:.1f}"
    s = str(v).strip().replace("／", "/")
    if s.endswith("%"):
        try:
            return f"{float(s[:-1]):.1f}"
        except ValueError:
            return s
    return s


def empty(v):
    return v is None or str(v).strip() == ""


def read_old(ws, n_fixed, ncols):
    out = {}
    last = [""] * n_fixed
    for row in ws.iter_rows(min_row=3, values_only=True):
        if any("平均" in str(c) for c in row if c is not None):
            break
        cells = list(row) + [None] * (n_fixed + ncols)
        fixed = [squash(str(c)) if c is not None else "" for c in cells[:n_fixed]]
        vals = cells[n_fixed:n_fixed + ncols]
        if not any(fixed) or all(empty(v) for v in vals):
            continue
        for i in range(n_fixed):
            if fixed[i]:
                last[i] = fixed[i]
            else:
                fixed[i] = last[i]
        if n_fixed == 3:
            fixed[2] = bonus_type(fixed[2])
        out.setdefault(tuple(fixed), [norm_val(v) for v in vals])
    return out


def index_new(rows, key_idx, n_fixed):
    out = {}
    for row in rows:
        key = tuple(squash(str(row[i])) for i in key_idx)
        out.setdefault(key, [norm_val(v) for v in row[n_fixed:]])
    return out


def compare(label, old, new):
    same = [k for k in new if k in old and old[k] == new[k]]
    diff = [(k, k) for k in new if k in old and old[k] != new[k]]
    only_new = [k for k in new if k not in old]
    only_old = [k for k in old if k not in new]
    used, renamed_same, left_new = set(), [], []
    for k in only_new:
        best, score = None, 0.0
        for o in only_old:
            if o in used or o[1:] != k[1:]:
                continue
            s = difflib.SequenceMatcher(None, k[0], o[0]).ratio()
            if s > score:
                best, score = o, s
        if best is not None and score >= 0.6:
            used.add(best)
            if new[k] == old[best]:
                renamed_same.append((k, best))
            else:
                diff.append((k, best))
        else:
            left_new.append(k)
    left_old = [o for o in only_old if o not in used]
    print(f"【{label}】新檔 {len(new)} 行｜舊檔 {len(old)} 行")
    print(f"  完全相同 {len(same)}｜名稱不同但數值相同 {len(renamed_same)}｜數值不同 {len(diff)}"
          f"｜只在新檔 {len(left_new)}｜只在舊檔 {len(left_old)}")
    for k, o in diff[:8]:
        print("  不同：", "/".join(k)[:60])
        print("    舊：", " ".join(old[o]))
        print("    新：", " ".join(new[k]))
    for k in left_new[:8]:
        print("  只在新檔：", "/".join(k)[:60])
    for k in left_old[:8]:
        print("  只在舊檔：", "/".join(k)[:60])


def to_frac(v):
    if is_num(v):
        return float(v)
    if isinstance(v, str) and v.strip().endswith("%"):
        try:
            return float(v.strip()[:-1]) / 100
        except ValueError:
            return None
    return None


def fold(s):
    return squash(str(s)).replace("紅", "红")


def old_summary(ws, col, ncols):
    section, out = None, {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        cells = list(row) + [None] * (col + ncols + 2)
        label = fold(cells[col]) if cells[col] is not None else ""
        if label == "分红類型":
            section = "cum"
        elif label == "當年平均值":
            section = "year"
        elif section and label:
            out[(section, label)] = [to_frac(c) for c in cells[col + 1:col + 1 + ncols]]
    return out


def old_cash_summary(ws, ncols, label):
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    for i, r in enumerate(rows):
        for j, c in enumerate(r):
            if c is not None and squash(str(c)) == "所有年份平均值" and i + 3 < len(rows):
                def pick(rr):
                    return [to_frac(x) for x in (rows[rr] + [None] * (j + ncols))[j:j + ncols]]
                return {("cum", fold(label)): pick(i + 1), ("year", fold(label)): pick(i + 3)}
    return {}


def check_summary(old, summary):
    mine = [("cum", line) for line in summary["cum"]] + [("year", line) for line in summary["year"]]
    for section, line in mine:
        name = ("累計 " if section == "cum" else "當年 ") + line[0]
        vals = line[1:]
        o = old.get((section, fold(line[0])))
        if o is None:
            print(f"  {name}：舊檔沒有這一行")
            continue
        ok = 0
        for a, b in zip(o, vals):
            if is_num(a) and is_num(b):
                ok += abs(a - b) < 0.00006
            elif not is_num(a) and not is_num(b):
                ok += 1
        print(f"  {name}：{ok}/{len(vals)} 格相同")
        if ok < len(vals):
            print("    舊：", " ".join(fmt_pct(x) for x in o))
            print("    新：", " ".join(fmt_pct(x) for x in vals))


# ---------- 主流程 ----------
c_recs = load_page(URL_CASH, "總現金價值比率")
f_recs = load_page(URL_FULFIL, "分紅實現率")
year = report_year(c_recs)
if not year or report_year(f_recs) != year:
    raise SystemExit(f"兩個頁面的報告年度不同或讀不到：{year} / {report_year(f_recs)}，暫不更新")

c_cols = max((rec["labels"] for rec in c_recs), key=len)
f_cols = max((rec["labels"] for rec in f_recs), key=len)
c_rows, c_skipped = build_rows(c_recs, c_cols, False, lambda lb: (not lb) or "總現金價值" in lb)
f_rows, _ = build_rows(f_recs, f_cols, True)

print()
print("=== 讀取結果 ===")
print(f"報告年度 {year}｜總現金價值比率 {len(c_rows)} 行｜分紅實現率 {len(f_rows)} 行")
print("年度欄：", " ".join(c_cols))
print("產品種類：", dict(collections.Counter(r[1] or "（空白）" for r in c_rows)))
types = collections.Counter(r[3] for r in f_rows)
print("紅利類型：", "｜".join(f"{k} {v}" for k, v in types.items()))
if c_skipped:
    print(f"總現金價值頁略過的其他表格 {len(c_skipped)} 個：", "；".join(c_skipped[:12]))
if len(c_cols) < 5 or len(f_cols) < 5 or len(c_rows) < 40 or len(f_rows) < 80:
    raise SystemExit("讀到的資料太少，保誠網站可能改了版面，暫不更新")

order = [t for t in PREFERRED if t in types] + [t for t in types if t not in PREFERRED]
f_groups = [("所有分紅", f_rows)]
for t in order:
    sub = [r for r in f_rows if r[3] == t]
    if any(is_num(v) for r in sub for v in r[4:]):
        f_groups.append((t, sub))
f_summary = build_summary(4, f_cols, f_groups)
c_summary = build_summary(3, c_cols, [("所有產品", c_rows)])

today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d")
legend = ("「/」＝沒有數字（產品當時尚未推出、已停售，或網站顯示 N/A）。紅色字＝低於 80%。\n"
          f"資料來源：保誠香港網站（報告年度 {year}）｜更新日期：{today}")
c_notes = ("總現金價值比率＝實際總現金價值（包括保證現金價值及非保證元素）÷ 銷售時之參考金額。"
           "100% 即與銷售時一致，高於 100% 即比銷售時多。\n" + legend)
f_notes = ("分紅實現率＝實際派發的累積非保證利益 ÷ 銷售時預期的金額。"
           "100% 即與預期一致，高於 100% 即比預期多。\n" + legend)

wb = Workbook()
wb.remove(wb.active)
add_sheet(wb, f"{year}總現金價值比率", f"保誠總現金價值比率（報告年度 {year}）", c_notes,
          make_header(c_cols, ["产品名称", "产品种类", "货币"]), c_rows, c_summary,
          "總現金價值比率", 3, [44, 16, 10])
add_sheet(wb, f"{year}分紅實現率", f"保誠分紅實現率（報告年度 {year}）", f_notes,
          make_header(f_cols, ["产品名称", "产品种类", "货币", "红利类型"]), f_rows, f_summary,
          "分紅類型", 4, [44, 16, 10, 22])
buf = io.BytesIO()
wb.save(buf)
content = buf.getvalue()
print(f"Excel 檔案 {len(content) // 1024} KB")

source = resolve_source(TARGET)
pages = query_pages(source)
page = next((p for p in pages if str(year) in page_title(p)), None)

print()
print("=== 與手動整理的檔案核對 ===")
old_wb = load_manual_file(page["id"]) if page else None
if old_wb is None:
    print(f"沒有 {year} 年的手動檔案，略過核對")
else:
    old_f = next((s for s in old_wb.worksheets if "分紅" in s.title or "分红" in s.title), None)
    old_c = next((s for s in old_wb.worksheets if "現金" in s.title or "现金" in s.title), None)
    if old_c is not None:
        compare("總現金價值比率", read_old(old_c, 2, len(c_cols)), index_new(c_rows, (0, 2), 3))
        print("【總現金價值比率平均值】")
        check_summary(old_cash_summary(old_c, len(c_cols), "所有產品"), c_summary)
    if old_f is not None:
        compare("分紅實現率", read_old(old_f, 3, len(f_cols)), index_new(f_rows, (0, 2, 3), 4))
        print("【分紅實現率平均值】")
        check_summary(old_summary(old_f, 2, len(f_cols)), f_summary)

print()
print("=== 上傳到 Notion ===")
if page is None:
    body = {"parent": {"type": "data_source_id", "data_source_id": source},
            "properties": {"title": {"title": [{"text": {"content": str(year)}}]}}}
    r = requests.post(f"{API}/pages", headers=HEADERS, json=body, timeout=30)
    page = check(r, "建立年度頁面")
    print(f"已建立頁面：{year}")
else:
    print("使用現有頁面：", page_title(page))
upload_file(page["id"], f"保誠總現金價值&分紅實現率（{year}年）{AUTO}.xlsx", content)
print("完成")

import html as htmllib
import re

import requests

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
URLS = [
    "https://www.prudential.com.hk/tc/general/performance/total-cash-value-ratio/",
    "https://www.prudential.com.hk/tc/general/performance/fulfillment-ratio/",
]


def text(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", "", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def attr(tag, name):
    m = re.search(r"\s" + name + r'="([^"]*)"', tag)
    return m.group(1) if m else ""


for url in URLS:
    print()
    print("=====", url.rstrip("/").split("/")[-1])
    r = requests.get(url, headers={"User-Agent": UA, "Accept-Language": "zh-HK,zh;q=0.9"}, timeout=60)
    r.encoding = "utf-8"
    page = r.text
    print(f"狀態 {r.status_code}｜長度 {len(page)}")

    tabs = [(re.sub(r"-tab$", "", attr(m.group(1), "id")), text(m.group(2)))
            for m in re.finditer(r'(<li\b[^>]*role="tab"[^>]*>)(.*?)</li>', page, flags=re.S)]
    print(f"分頁 {len(tabs)} 個：", "｜".join(f"{i}={n}" for i, n in tabs))

    panels = [(m.start(), attr(m.group(0), "id"))
              for m in re.finditer(r'<div\b[^>]*role="tabpanel"[^>]*>', page)]
    tables = [(m.start(), m.group(0))
              for m in re.finditer(r"<table\b.*?</table>", page, flags=re.S | re.I)]

    def owner(pos):
        pid = "(無)"
        for start, i in panels:
            if start > pos:
                break
            pid = i
        return pid

    groups = {}
    for pos, t in tables:
        groups.setdefault(owner(pos), []).append((pos, t))
    print(f"表格 {len(tables)} 個：", " ".join(f"{k.replace('-tabpanel', '')}:{len(v)}" for k, v in groups.items()))

    picks = list(groups)[:1]
    if "egs-tabpanel" in groups and "egs-tabpanel" not in picks:
        picks.append("egs-tabpanel")
    for pid in picks:
        print(f"--- 面板 {pid}")
        for n, (pos, t) in enumerate(groups[pid][:2], 1):
            rows = re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I)
            print(f"[表格{n}｜{len(rows)} 行] 前文：{text(page[max(0, pos - 1500):pos])[-160:]}")
            for tr in rows[:8]:
                cells = []
                for tag, at, body in re.findall(r"<(t[hd])\b([^>]*)>(.*?)</t[hd]>", tr, flags=re.S | re.I):
                    mark = "".join(f"[{k[0]}{attr(' ' + at, k)}]" for k in ("colspan", "rowspan") if attr(" " + at, k))
                    cells.append(text(body) + mark)
                print("  " + " | ".join(cells))

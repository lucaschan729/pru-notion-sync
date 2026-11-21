import re

import requests

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
URLS = [
    "https://www.prudential.com.hk/tc/general/performance/total-cash-value-ratio/",
    "https://www.prudential.com.hk/tc/general/performance/fulfillment-ratio/",
    "https://www.prudential.com.hk/performance/fulfillment-ratio/tc/",
]

for url in URLS:
    print()
    print("=====", url)
    try:
        r = requests.get(url, headers={"User-Agent": UA, "Accept-Language": "zh-HK,zh;q=0.9"}, timeout=60)
    except Exception as e:
        print("失敗：", str(e)[:200])
        continue
    r.encoding = "utf-8"
    html = r.text
    print(f"狀態 {r.status_code}｜最終網址 {r.url}｜長度 {len(html)}")
    print(f"table {html.count('<table')}｜% {html.count('%')}｜雋陞 {html.count('雋陞')}｜iframe {html.count('<iframe')}")
    files = sorted(set(re.findall(r"[\"']([^\"'\s<>]+\.(?:json|xlsx|csv|pdf)[^\"'\s<>]*)[\"']", html)))
    for x in files[:40]:
        print("  檔案：", x[:200])
    apis = sorted(set(re.findall(r"[\"']([^\"'\s<>]*(?:/api/|graphql|ajax|/bin/)[^\"'\s<>]*)[\"']", html)))
    for x in apis[:40]:
        print("  介面：", x[:200])
    frames = re.findall(r"<iframe[^>]+src=[\"']([^\"']+)", html)
    for x in frames[:10]:
        print("  iframe：", x[:200])
    i = html.find("雋陞")
    if i >= 0:
        print("--- 產品名稱附近的內容 ---")
        print(re.sub(r"\s+", " ", html[max(0, i - 300):i + 1500]))

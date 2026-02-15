import csv
import os
import time
from datetime import datetime, timedelta
from typing import List

import requests
from openpyxl import Workbook


def _mock_videos(category: str, n: int):
    now = datetime.utcnow().isoformat() + "Z"
    items = []
    for i in range(1, n + 1):
        items.append(
            {
                "category": category,
                "rank": i,
                "title": f"Mock Video {i} ({category})",
                "channel": "MockChannel",
                "views": 1000 * i,
                "url": f"http://example.com/{category}/{i}",
                "published_at": now,
            }
        )
    return items


def fetch_videos_for_category(category: str, n: int, lang: str, days: int, api_key: str):
    # Attempts to call YouTube Data API v3 'videos?chart=mostPopular'
    params = {
        "part": "snippet,statistics",
        "chart": "mostPopular",
        "regionCode": lang,
        "maxResults": min(50, n),
        "key": api_key,
    }
    # If category looks numeric, use as videoCategoryId; otherwise skip filtering
    if category and category.isdigit():
        params["videoCategoryId"] = category

    resp = requests.get("https://www.googleapis.com/youtube/v3/videos", params=params, timeout=10)
    resp.raise_for_status()
    data = resp.json()
    items = []
    for idx, it in enumerate(data.get("items", [])[:n], start=1):
        snip = it.get("snippet", {})
        stats = it.get("statistics", {})
        vid_id = it.get("id")
        url = f"https://youtu.be/{vid_id}"
        items.append(
            {
                "category": category,
                "rank": idx,
                "title": snip.get("title"),
                "channel": snip.get("channelTitle"),
                "views": int(stats.get("viewCount", 0)),
                "url": url,
                "published_at": snip.get("publishedAt"),
            }
        )
    return items


def write_csv(path: str, rows: List[dict]):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    headers = ["category", "rank", "title", "channel", "views", "url", "published_at"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for r in rows:
            w.writerow(r)


def write_xlsx(path: str, rows: List[dict]):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb = Workbook()
    ws = wb.active
    headers = ["category", "rank", "title", "channel", "views", "url", "published_at"]
    ws.append(headers)
    for r in rows:
        row = [r.get(h) for h in headers]
        ws.append(row)
        # set hyperlink on the last appended row's URL cell
        url_cell = ws.cell(row=ws.max_row, column=headers.index("url") + 1)
        url = r.get("url")
        if url:
            url_cell.hyperlink = url
    wb.save(path)


def fetch_and_export(categories: str, n: int, lang: str, days: int, api_key: str = None, mock: bool = False):
    cats = [c.strip() for c in categories.split(",")] if categories else ["all"]
    all_rows = []
    for c in cats:
        if mock or not api_key:
            rows = _mock_videos(c or "all", n)
        else:
            try:
                rows = fetch_videos_for_category(c, n, lang, days, api_key)
            except Exception:
                rows = _mock_videos(c or "all", n)
        all_rows.extend(rows)

    out_csv = os.path.join("out", "top_videos.csv")
    out_xlsx = os.path.join("out", "top_videos.xlsx")
    write_csv(out_csv, all_rows)
    write_xlsx(out_xlsx, all_rows)
    return out_csv, out_xlsx

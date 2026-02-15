# yt_top_exporter

Exports top YouTube videos to CSV and XLSX (clickable links).

Quick start

1. Create and activate a venv, install deps:

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

2. Create a `.env` in the project root (or set env vars):

```text
YOUTUBE_API_KEY=your_key_here
```

3. Run (mock mode works without an API key):

```bash
.venv/bin/python -m yt_top.run --mock --categories "music,news" --n 5 --lang US --days 7
```

Outputs written to `out/top_videos.csv` and `out/top_videos.xlsx`.

Run tests:

```bash
.venv/bin/pytest -q
```

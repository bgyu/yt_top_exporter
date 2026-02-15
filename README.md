# yt_top_exporter

Exports top YouTube videos to CSV and XLSX (clickable links).

## Requirements

- Python 3.11+
- Install required packages with:

```bash
python -m pip install -r requirements.txt
```

Optional: `pandas` and `openpyxl` for a more robust CSV→XLSX conversion.

## Quick start

1. Create and activate a virtualenv, then install deps:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2. Provide a YouTube API key either via environment or a `.env` file at the project root:

```text
YOUTUBE_API_KEY=your_key_here
```

3. Run the exporter (real data — API key required):

```bash
python -m yt_top.run --categories all --n 5 --lang US --days 30
```

For testing without an API key use explicit mock mode:

```bash
python -m yt_top.run --mock
```

Outputs are written to the `out/` directory:

- `out/top_videos.csv` — raw CSV
- `out/top_videos.xlsx` — XLSX with hyperlinks
- `out/youtube_top_videos_last_{days}_{lang}.csv` — enriched CSV

## CSV → XLSX conversion

If you have `pandas` and `openpyxl` installed you can create a robust XLSX from an enriched CSV:

```bash
pip install pandas openpyxl
python scripts/csv_to_xlsx_pandas.py out/youtube_top_videos_last_30_US.csv
```

If those libraries are not available, use the builtin converter (no extra deps):

```bash
python scripts/csv_to_xlsx_builtin.py out/youtube_top_videos_last_30_US.csv
```

## Notes

- The CLI requires `YOUTUBE_API_KEY` unless you pass `--mock` explicitly.
- When fetching real data, some categories may be skipped if the YouTube API returns an error for that category; the exporter will log and continue.
- If Excel reports an `.xlsx` as corrupted, convert the enriched CSV with `pandas`/`openpyxl` on a machine that has those packages installed.

## Tests

Run tests with:

```bash
pytest -q
```

Feel free to ask me to run the real export now or to update any of these instructions.

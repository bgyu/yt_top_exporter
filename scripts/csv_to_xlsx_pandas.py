import sys
from pathlib import Path


def main():
    # allow user to pass CSV path; otherwise pick the first matching enriched CSV in out/
    if len(sys.argv) > 1:
        csv_path = Path(sys.argv[1])
    else:
        # prefer CSV candidates first
        candidates = list(Path('out').glob('youtube_top_videos_last_*.csv'))
        if not candidates:
            # fall back to any matching file if no CSVs found
            candidates = list(Path('out').glob('youtube_top_videos_last_*'))
        if not candidates:
            print('No matching CSV found in out/; provide a path argument')
            sys.exit(1)
        # prefer .csv
        csv_path = next((p for p in candidates if p.suffix.lower() == '.csv'), candidates[0])

    if not csv_path.exists():
        print(f'CSV not found: {csv_path}')
        sys.exit(1)

    xlsx_path = csv_path.with_suffix('.xlsx')

    # Try using pandas (recommended). If not available, fall back to the builtin writer.
    try:
        import pandas as pd

        print(f'Reading {csv_path}...')
        df = pd.read_csv(csv_path)
        # normalize column name: some enriched CSVs use `video_url`; rename to `url`
        if 'video_url' in df.columns and 'url' not in df.columns:
            df = df.rename(columns={'video_url': 'url'})
        print(f'Writing {xlsx_path}...')
        df.to_excel(xlsx_path, index=False)
        print(f'Done: {xlsx_path}')
        return
    except Exception as e:
        # pandas not available or failed; fall back
        print('pandas not available or failed:', e)

    # Fallback: use the repository minimal XLSX writer
    try:
        # ensure repo root on path
        sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
        from yt_top.exporter import write_xlsx_minimal
        import csv

        print('Falling back to builtin XLSX writer')
        with open(csv_path, newline='', encoding='utf-8') as f:
            r = csv.DictReader(f)
            rows = [row for row in r]

        headers = ["category", "rank", "title", "channel", "views", "url", "published_at"]
        norm_rows = []
        for row in rows:
            norm = {h: row.get(h, '') for h in headers}
            norm_rows.append(norm)

        write_xlsx_minimal(str(xlsx_path), norm_rows)
        print('Wrote (builtin):', xlsx_path)
    except Exception as e:
        print('Failed to write XLSX:', e)
        sys.exit(1)


if __name__ == '__main__':
    main()

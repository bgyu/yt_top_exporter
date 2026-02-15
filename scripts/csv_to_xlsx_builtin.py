import sys
import csv
from pathlib import Path

# ensure repository root is on sys.path so `yt_top` package imports work when running from /scripts
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

def main():
    if len(sys.argv) < 2:
        print('Usage: csv_to_xlsx_builtin.py <csv-path> [xlsx-path]')
        sys.exit(1)
    csv_path = Path(sys.argv[1])
    if not csv_path.exists():
        print(f'CSV not found: {csv_path}')
        sys.exit(1)
    xlsx_path = Path(sys.argv[2]) if len(sys.argv) > 2 else csv_path.with_suffix('.xlsx')

    # read CSV rows into list of dicts
    with open(csv_path, newline='', encoding='utf-8') as f:
        r = csv.DictReader(f)
        rows = [row for row in r]

    # Normalize rows to expected keys for exporter.write_xlsx_minimal
    headers = ["category", "rank", "title", "channel", "views", "url", "published_at"]
    norm_rows = []
    for row in rows:
        norm = {h: row.get(h, '') for h in headers}
        norm_rows.append(norm)

    # import exporter writer and write xlsx
    from yt_top.exporter import write_xlsx_minimal

    write_xlsx_minimal(str(xlsx_path), norm_rows)
    print('Wrote:', xlsx_path)

if __name__ == '__main__':
    main()

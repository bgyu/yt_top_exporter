import argparse
import os
from dotenv import load_dotenv
from . import exporter


def build_parser():
    p = argparse.ArgumentParser(description="YouTube Top Videos exporter")
    p.add_argument("--categories", default="all", help="Comma-separated categories (names or ids)")
    p.add_argument("--n", type=int, default=5, help="Top N per category")
    p.add_argument("--lang", default="US", help="Region/language code (used as regionCode)")
    p.add_argument("--days", type=int, default=7, help="Time window in days (informational)")
    p.add_argument("--mock", action="store_true", help="Run in mock mode (no API calls)")
    return p


def main(argv=None):
    load_dotenv()
    parser = build_parser()
    args = parser.parse_args(argv)

    api_key = os.getenv("YOUTUBE_API_KEY")
    mock = args.mock

    if not mock and not api_key:
        parser.error("YOUTUBE_API_KEY not set in environment. Set it or run with --mock.")

    results = exporter.fetch_and_export(
        categories=args.categories,
        n=args.n,
        lang=args.lang,
        days=args.days,
        api_key=api_key,
        mock=mock,
    )
    # fetch_and_export may return (csv, xlsx) or (csv, xlsx, enriched_csv)
    if isinstance(results, tuple) or isinstance(results, list):
        print("Wrote:", *results)
    else:
        print("Wrote:", results)


if __name__ == "__main__":
    main()

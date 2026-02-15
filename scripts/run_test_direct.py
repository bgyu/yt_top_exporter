import tempfile
import pathlib

import sys
import pathlib as _pathlib
sys.path.insert(0, str(_pathlib.Path(__file__).resolve().parents[1]))
import tests.test_end_to_end as t


def main():
    with tempfile.TemporaryDirectory() as d:
        p = pathlib.Path(d)
        t.test_end_to_end_mock(p)


if __name__ == "__main__":
    main()

import os
import sys
import pathlib

# Ensure project root is on sys.path so imports work under pytest
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))

from yt_top import run, verifier


def test_end_to_end_mock(tmp_path):
    # ensure out dir is isolated to tmp_path by switching cwd
    cwd = os.getcwd()
    try:
        os.chdir(tmp_path)
        # run in mock mode, creates out/top_videos.csv and xlsx
        run.main(["--mock", "--categories", "testcat", "--n", "1"])
        csvp = os.path.join("out", "top_videos.csv")
        xlsxp = os.path.join("out", "top_videos.xlsx")
        assert os.path.exists(csvp)
        assert os.path.exists(xlsxp)
        assert verifier.verify_all(csvp, xlsxp)
    finally:
        os.chdir(cwd)

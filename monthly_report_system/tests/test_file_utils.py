from __future__ import annotations

import os
import time
from pathlib import Path

from utils import file_utils
from utils.file_utils import cleanup_old_run_dirs, cleanup_run_dir, safe_upload_filename


def test_safe_upload_filename_strips_paths_and_keeps_chinese():
    assert safe_upload_filename(r"..\资料/五月检查表.xlsx", "data.xlsx") == "五月检查表.xlsx"
    assert safe_upload_filename("bad:name?.docx", "template.docx") == "bad_name_.docx"


def test_cleanup_old_run_dirs_only_removes_expired_run_dirs(tmp_path, monkeypatch):
    old_run = tmp_path / "run_old"
    fresh_run = tmp_path / "run_fresh"
    other_dir = tmp_path / "keep_me"
    old_run.mkdir()
    fresh_run.mkdir()
    other_dir.mkdir()

    old_time = time.time() - 10 * 24 * 60 * 60
    os.utime(old_run, (old_time, old_time))
    monkeypatch.setattr(file_utils, "TEMP_DIR", Path(tmp_path))

    assert cleanup_old_run_dirs(retention_days=7) == []
    assert not old_run.exists()
    assert fresh_run.exists()
    assert other_dir.exists()


def test_cleanup_run_dir_only_removes_run_dir_under_temp(tmp_path, monkeypatch):
    run_dir = tmp_path / "run_current"
    other_dir = tmp_path / "other"
    run_dir.mkdir()
    other_dir.mkdir()
    monkeypatch.setattr(file_utils, "TEMP_DIR", Path(tmp_path))

    assert cleanup_run_dir(run_dir) is None
    assert not run_dir.exists()
    assert cleanup_run_dir(other_dir) is not None
    assert other_dir.exists()

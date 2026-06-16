"""文件和目录工具。"""

from __future__ import annotations

from datetime import datetime, timedelta
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import BinaryIO


if getattr(sys, "frozen", False):
    PROJECT_ROOT = Path(sys.executable).resolve().parent
else:
    PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEMP_DIR = PROJECT_ROOT / "temp"
OUTPUT_DIR = PROJECT_ROOT / "output"
RUN_DIR_RETENTION_DAYS = 7


def ensure_dirs() -> None:
    """确保运行所需目录存在。"""

    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def safe_filename(filename: str, default: str = "月报.docx") -> str:
    """清理 Windows 不允许的输出文件名字符。"""

    name = (filename or default).strip()
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    if not name.lower().endswith(".docx"):
        name += ".docx"
    return name or default


def safe_upload_filename(filename: str, fallback: str) -> str:
    """清理输入文件名，去掉路径和 Windows 非法字符，同时保留中文。"""

    original = (filename or fallback).strip().replace("\\", "/")
    name = original.rsplit("/", 1)[-1].strip() or fallback
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    return name.strip(". ") or fallback


def cleanup_old_run_dirs(retention_days: int = RUN_DIR_RETENTION_DAYS) -> list[str]:
    """清理过期 run_* 临时目录，返回无法清理的目录说明。"""

    ensure_dirs()
    cutoff = datetime.now() - timedelta(days=retention_days)
    warnings: list[str] = []
    for path in TEMP_DIR.glob("run_*"):
        try:
            if not path.is_dir():
                continue
            modified_at = datetime.fromtimestamp(path.stat().st_mtime)
            if modified_at >= cutoff:
                continue
            shutil.rmtree(path)
        except PermissionError:
            warnings.append(f"临时目录被占用或无权限清理，已跳过：{path.name}")
        except OSError as exc:
            warnings.append(f"临时目录清理失败，已跳过：{path.name}（{exc}）")
    return warnings


def cleanup_run_dir(path: Path) -> str | None:
    """清理单次运行目录，成功返回 None，失败返回提示。"""

    try:
        resolved = path.resolve()
        temp_root = TEMP_DIR.resolve()
        if temp_root not in resolved.parents or not resolved.name.startswith("run_"):
            return f"临时目录清理跳过：{resolved} 不在允许范围内。"
        if resolved.exists():
            shutil.rmtree(resolved)
    except PermissionError:
        return f"临时目录被占用或无权限清理，已跳过：{path.name}"
    except OSError as exc:
        return f"临时目录清理失败，已跳过：{path.name}（{exc}）"
    return None


def create_run_dir(prefix: str = "run_") -> Path:
    """创建一次处理过程专用临时目录。"""

    ensure_dirs()
    cleanup_old_run_dirs()
    return Path(tempfile.mkdtemp(prefix=prefix, dir=TEMP_DIR))


def save_upload(uploaded_file: BinaryIO, target_dir: Path, filename: str) -> Path:
    """保存类文件对象到指定目录。"""

    target_dir.mkdir(parents=True, exist_ok=True)
    path = target_dir / filename
    with path.open("wb") as f:
        shutil.copyfileobj(uploaded_file, f)
    return path

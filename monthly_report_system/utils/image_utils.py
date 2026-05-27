"""图片处理工具。"""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Tuple

from PIL import Image, ImageOps


def save_normalized_image(
    img: Image.Image,
    output_path: Path,
    quality: int = 72,
    target_size: Tuple[int, int] = (900, 525),
) -> Path:
    """把已打开的图片对象保存为 Word 友好的 JPEG。"""

    img.draft("RGB", target_size)
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    if img.mode == "L":
        img = img.convert("RGB")

    canvas = Image.new("RGB", target_size, "white")
    img.thumbnail(target_size, Image.Resampling.LANCZOS, reducing_gap=3.0)
    left = (target_size[0] - img.width) // 2
    top = (target_size[1] - img.height) // 2
    canvas.paste(img, (left, top))
    canvas.save(output_path, format="JPEG", quality=quality, optimize=True)
    return output_path


def normalize_image(
    image_path: Path,
    output_path: Optional[Path] = None,
    quality: int = 72,
    target_size: Tuple[int, int] = (900, 525),
) -> Path:
    """把图片转为适合插入 Word 的 JPEG。

    Word 中只需要按 10.12cm x 5.72cm 显示，直接嵌入原始大图会让
    docx 体积暴涨。这里将图片等比缩放到目标画布中，不拉伸变形。
    """

    output_path = output_path or image_path.with_suffix(".jpg")
    with Image.open(image_path) as img:
        return save_normalized_image(img, output_path, quality=quality, target_size=target_size)

"""图片处理工具。"""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Tuple

from PIL import Image, ImageOps


def _clamp_box(box: tuple[int, int, int, int], width: int, height: int) -> tuple[int, int, int, int]:
    left, top, right, bottom = box
    return max(0, left), max(0, top), min(width, right), min(height, bottom)


def _expand_box_to_aspect(
    box: tuple[int, int, int, int],
    image_size: tuple[int, int],
    target_aspect: float,
) -> tuple[int, int, int, int]:
    """扩展裁剪框到目标比例，只扩大不压缩，避免裁掉有效内容。"""

    width, height = image_size
    left, top, right, bottom = box
    crop_width = max(1, right - left)
    crop_height = max(1, bottom - top)
    current_aspect = crop_width / crop_height

    if abs(current_aspect - target_aspect) < 0.01:
        return _clamp_box(box, width, height)

    if current_aspect < target_aspect:
        new_width = int(round(crop_height * target_aspect))
        delta = max(0, new_width - crop_width)
        left -= delta // 2
        right += delta - delta // 2
        if left < 0:
            right -= left
            left = 0
        if right > width:
            left -= right - width
            right = width
    else:
        new_height = int(round(crop_width / target_aspect))
        delta = max(0, new_height - crop_height)
        top -= delta // 2
        bottom += delta - delta // 2
        if top < 0:
            bottom -= top
            top = 0
        if bottom > height:
            top -= bottom - height
            bottom = height

    return _clamp_box((left, top, right, bottom), width, height)


def _trim_white_border(
    img: Image.Image,
    target_size: Tuple[int, int],
    white_threshold: int = 245,
    margin: int = 4,
) -> Image.Image:
    """裁掉图片四周近白边，并把裁剪框扩展回目标比例。

    只裁连续白边；内容区域的比例按目标画布比例扩展，避免后续插入 Word 时变形。
    """

    rgb = img.convert("RGB")
    mask = rgb.point(lambda value: 255 if value < white_threshold else 0).convert("L")
    bbox = mask.getbbox()
    if bbox is None:
        return rgb

    width, height = rgb.size
    left, top, right, bottom = bbox
    box = _clamp_box((left - margin, top - margin, right + margin, bottom + margin), width, height)

    # 没有明显白边时不处理，避免误裁浅色天空、墙面等真实内容。
    if box == (0, 0, width, height):
        return rgb

    target_aspect = target_size[0] / target_size[1]
    box = _expand_box_to_aspect(box, rgb.size, target_aspect)
    if box == (0, 0, width, height):
        return rgb
    return rgb.crop(box)


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
    img = _trim_white_border(img, target_size)

    canvas = Image.new("RGB", target_size, "white")
    img = ImageOps.contain(img, target_size, Image.Resampling.LANCZOS)
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

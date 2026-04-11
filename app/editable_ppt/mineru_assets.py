from __future__ import annotations

import io
import json
import math
import time
import uuid
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

import requests

from .assets import read_image_size
from .browser import ensure_dir

MINERU_DEFAULT_BASE_URL = "https://mineru.net/api/v4"
MINERU_MAX_UPLOAD_BYTES = 10 * 1000 * 1000
MINERU_MIN_UPLOAD_EDGE = 720
MINERU_UPLOAD_SCALE_STEPS = (1.0, 0.9, 0.8, 0.72, 0.64, 0.56, 0.48, 0.4, 0.33, 0.28, 0.24)
MINERU_UPLOAD_JPEG_QUALITIES = (90, 84, 78, 72, 66, 60)
VISUAL_ELEMENT_TYPES = {"image", "table", "equation"}
UNMATCHED_PH_REGION_EXPAND_RATIO = 1.5
MATCH_MIN_ASPECT_SCORE = 0.25
MATCH_MIN_AREA_SCORE = 0.18
MATCH_MIN_IOU = 0.12
MATCH_MIN_PLACEHOLDER_COVER = 0.45
MATCH_MAX_ELEMENT_TO_PLACEHOLDER_RATIO = 8.5
MATCH_OVERSIZED_RATIO = 1.75
MATCH_OVERSIZED_RATIO_HARD = 2.25
MATCH_LOW_ELEMENT_COVER = 0.58
MATCH_LOW_IOU_FOR_OVERSIZED = 0.4
PLACEHOLDER_CUTOUT_EXPAND_RATIO = 0.12
PLACEHOLDER_CUTOUT_FALLBACK_EXPAND_RATIO = 0.2
PLACEHOLDER_CUTOUT_INSET_X_RATIO = 0.04
PLACEHOLDER_CUTOUT_INSET_TOP_RATIO = 0.01
PLACEHOLDER_CUTOUT_INSET_BOTTOM_RATIO = 0.05
SCREENSHOT_CUTOUT_INSET_X_RATIO = 0.08
SCREENSHOT_CUTOUT_INSET_TOP_RATIO = 0.015
SCREENSHOT_CUTOUT_INSET_BOTTOM_RATIO = 0.09
SCREENSHOT_SUBJECT_SAFE_X_RATIO = 0.03
SCREENSHOT_SUBJECT_SAFE_TOP_RATIO = 0.01
SCREENSHOT_SUBJECT_SAFE_BOTTOM_RATIO = 0.08
SCREENSHOT_SUBJECT_PAD_X_RATIO = 0.035
SCREENSHOT_SUBJECT_PAD_TOP_RATIO = 0.025
SCREENSHOT_SUBJECT_PAD_BOTTOM_RATIO = 0.03
GRAPHIC_CUTOUT_BLEED_RATIO = 0.05
SALVAGE_MIN_PLACEHOLDER_COVER = 0.82
SALVAGE_MAX_CENTER_DISTANCE = 0.08


@dataclass(frozen=True)
class MineruPreparedUpload:
    file_path: Path
    original_size: tuple[int, int]
    upload_size: tuple[int, int]
    file_size_bytes: int
    mime_type: str


@dataclass(frozen=True)
class MineruParseResult:
    extracted_dir: Path
    prepared_upload: MineruPreparedUpload


def _load_pillow():
    try:
        from PIL import Image, ImageDraw
    except ImportError as exc:  # pragma: no cover - depends on runtime
        raise RuntimeError("MinerU asset extraction requires Pillow.") from exc
    return Image, ImageDraw


def _lanczos_resample(Image: Any) -> Any:
    resampling = getattr(Image, "Resampling", None)
    if resampling is not None:
        return resampling.LANCZOS
    return getattr(Image, "LANCZOS")


def _flatten_image_for_jpeg(image: Any) -> Any:
    if image.mode in {"RGB", "L"}:
        return image.convert("RGB")
    Image, _ = _load_pillow()
    rgba = image.convert("RGBA")
    background = Image.new("RGB", rgba.size, (255, 255, 255))
    background.paste(rgba, mask=rgba.getchannel("A"))
    return background


def _resize_to_fit(
    size: tuple[int, int],
    scale: float,
) -> tuple[int, int]:
    width, height = size
    scaled_width = max(1, int(round(width * scale)))
    scaled_height = max(1, int(round(height * scale)))
    if min(scaled_width, scaled_height) >= MINERU_MIN_UPLOAD_EDGE:
        return scaled_width, scaled_height

    min_side = max(min(width, height), 1)
    floor_scale = min(1.0, MINERU_MIN_UPLOAD_EDGE / min_side)
    return (
        max(1, int(round(width * floor_scale))),
        max(1, int(round(height * floor_scale))),
    )


def _encode_jpeg_bytes(image: Any, quality: int) -> bytes:
    buffer = io.BytesIO()
    image.save(
        buffer,
        format="JPEG",
        quality=int(quality),
        optimize=True,
        progressive=True,
    )
    return buffer.getvalue()


def _write_upload_metadata(
    metadata_path: Path,
    *,
    original_path: Path,
    prepared: MineruPreparedUpload,
) -> None:
    payload = {
        "original_path": str(original_path),
        "prepared_path": str(prepared.file_path),
        "original_size": list(prepared.original_size),
        "upload_size": list(prepared.upload_size),
        "file_size_bytes": prepared.file_size_bytes,
        "mime_type": prepared.mime_type,
        "compressed": prepared.file_path.resolve() != original_path.resolve(),
        "max_upload_bytes": MINERU_MAX_UPLOAD_BYTES,
    }
    metadata_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _prepare_image_for_mineru_upload(
    *,
    file_path: Path,
    work_dir: Path,
) -> MineruPreparedUpload:
    original_size = read_image_size(file_path)
    original_bytes = file_path.stat().st_size
    if original_bytes <= MINERU_MAX_UPLOAD_BYTES:
        prepared = MineruPreparedUpload(
            file_path=file_path,
            original_size=original_size,
            upload_size=original_size,
            file_size_bytes=original_bytes,
            mime_type="image/png" if file_path.suffix.lower() == ".png" else "image/jpeg",
        )
        _write_upload_metadata(work_dir / "prepared_upload.json", original_path=file_path, prepared=prepared)
        return prepared

    Image, _ = _load_pillow()
    upload_dir = work_dir / "prepared_upload"
    ensure_dir(upload_dir)

    with Image.open(file_path) as opened:
        source = _flatten_image_for_jpeg(opened)
        source.load()

    resample = _lanczos_resample(Image)
    best_bytes: bytes | None = None
    best_size: tuple[int, int] = source.size

    seen_sizes: set[tuple[int, int]] = set()
    for scale in MINERU_UPLOAD_SCALE_STEPS:
        candidate_size = _resize_to_fit(source.size, scale)
        if candidate_size in seen_sizes:
            continue
        seen_sizes.add(candidate_size)
        if candidate_size == source.size:
            resized = source
        else:
            resized = source.resize(candidate_size, resample=resample)
        for quality in MINERU_UPLOAD_JPEG_QUALITIES:
            encoded = _encode_jpeg_bytes(resized, quality)
            best_bytes = encoded
            best_size = candidate_size
            if len(encoded) <= MINERU_MAX_UPLOAD_BYTES:
                output_path = upload_dir / f"{file_path.stem}_edit_ready.jpg"
                output_path.write_bytes(encoded)
                prepared = MineruPreparedUpload(
                    file_path=output_path,
                    original_size=original_size,
                    upload_size=candidate_size,
                    file_size_bytes=len(encoded),
                    mime_type="image/jpeg",
                )
                _write_upload_metadata(work_dir / "prepared_upload.json", original_path=file_path, prepared=prepared)
                return prepared

    if not best_bytes:
        raise RuntimeError(f"Failed to prepare MinerU upload for {file_path}.")

    output_path = upload_dir / f"{file_path.stem}_edit_ready.jpg"
    output_path.write_bytes(best_bytes)
    prepared = MineruPreparedUpload(
        file_path=output_path,
        original_size=original_size,
        upload_size=best_size,
        file_size_bytes=len(best_bytes),
        mime_type="image/jpeg",
    )
    _write_upload_metadata(work_dir / "prepared_upload.json", original_path=file_path, prepared=prepared)
    return prepared


def _clip_int(value: Any, lower: int, upper: int) -> int:
    try:
        parsed = int(round(float(value)))
    except (TypeError, ValueError):
        parsed = lower
    return max(lower, min(parsed, upper))


def _rect_area(bbox: list[int]) -> int:
    return max(0, bbox[2] - bbox[0]) * max(0, bbox[3] - bbox[1])


def _intersection_area(a: list[int], b: list[int]) -> int:
    x1 = max(a[0], b[0])
    y1 = max(a[1], b[1])
    x2 = min(a[2], b[2])
    y2 = min(a[3], b[3])
    return max(0, x2 - x1) * max(0, y2 - y1)


def _intersect_bbox(a: list[int], b: list[int]) -> Optional[list[int]]:
    x1 = max(a[0], b[0])
    y1 = max(a[1], b[1])
    x2 = min(a[2], b[2])
    y2 = min(a[3], b[3])
    if x2 <= x1 or y2 <= y1:
        return None
    return [x1, y1, x2, y2]


def _iou(a: list[int], b: list[int]) -> float:
    inter = _intersection_area(a, b)
    if inter <= 0:
        return 0.0
    union = _rect_area(a) + _rect_area(b) - inter
    return inter / max(union, 1)


def _bbox_from_poly(poly: list[float], width: int, height: int) -> Optional[list[int]]:
    if len(poly) < 8:
        return None
    xs = [_clip_int(poly[index], 0, width) for index in range(0, len(poly), 2)]
    ys = [_clip_int(poly[index], 0, height) for index in range(1, len(poly), 2)]
    x1 = min(xs)
    y1 = min(ys)
    x2 = max(xs)
    y2 = max(ys)
    if x2 <= x1 or y2 <= y1:
        return None
    return [x1, y1, x2, y2]


def _bbox_from_any(raw_bbox: Any, width: int, height: int) -> Optional[list[int]]:
    if not isinstance(raw_bbox, list) or len(raw_bbox) != 4:
        return None
    try:
        values = [float(item) for item in raw_bbox]
    except (TypeError, ValueError):
        return None
    if max(values) <= 1.0001:
        scale_x = width
        scale_y = height
    elif max(values) <= 1000.0001:
        scale_x = width / 1000.0
        scale_y = height / 1000.0
    else:
        scale_x = 1.0
        scale_y = 1.0
    x1 = _clip_int(values[0] * scale_x, 0, max(width - 1, 0))
    y1 = _clip_int(values[1] * scale_y, 0, max(height - 1, 0))
    x2 = _clip_int(values[2] * scale_x, x1 + 1, width)
    y2 = _clip_int(values[3] * scale_y, y1 + 1, height)
    if x2 <= x1 or y2 <= y1:
        return None
    return [x1, y1, x2, y2]


def _remap_bbox_to_size(
    bbox: list[int],
    *,
    from_size: tuple[int, int],
    to_size: tuple[int, int],
) -> Optional[list[int]]:
    from_width = max(int(from_size[0]), 1)
    from_height = max(int(from_size[1]), 1)
    to_width = max(int(to_size[0]), 1)
    to_height = max(int(to_size[1]), 1)
    if (from_width, from_height) == (to_width, to_height):
        return list(bbox)

    scale_x = to_width / from_width
    scale_y = to_height / from_height
    x1 = _clip_int(bbox[0] * scale_x, 0, max(to_width - 1, 0))
    y1 = _clip_int(bbox[1] * scale_y, 0, max(to_height - 1, 0))
    x2 = _clip_int(bbox[2] * scale_x, x1 + 1, to_width)
    y2 = _clip_int(bbox[3] * scale_y, y1 + 1, to_height)
    if x2 <= x1 or y2 <= y1:
        return None
    return [x1, y1, x2, y2]


def _remap_elements_to_size(
    elements: list[dict[str, Any]],
    *,
    from_size: tuple[int, int],
    to_size: tuple[int, int],
) -> list[dict[str, Any]]:
    if tuple(from_size) == tuple(to_size):
        return [dict(item) for item in elements]

    remapped: list[dict[str, Any]] = []
    for item in elements:
        mapped_bbox = _remap_bbox_to_size(item["bbox"], from_size=from_size, to_size=to_size)
        if not mapped_bbox:
            continue
        mapped = dict(item)
        mapped["bbox_upload"] = list(item["bbox"])
        mapped["bbox"] = mapped_bbox
        remapped.append(mapped)
    return remapped


def _center_distance_norm(a: list[int], b: list[int], diagonal: float) -> float:
    ax = (a[0] + a[2]) / 2.0
    ay = (a[1] + a[3]) / 2.0
    bx = (b[0] + b[2]) / 2.0
    by = (b[1] + b[3]) / 2.0
    return math.hypot(ax - bx, ay - by) / max(diagonal, 1.0)


def _aspect_score(a: list[int], b: list[int]) -> float:
    aw = max(a[2] - a[0], 1)
    ah = max(a[3] - a[1], 1)
    bw = max(b[2] - b[0], 1)
    bh = max(b[3] - b[1], 1)
    ratio = math.log((aw / ah) / max(bw / bh, 1e-6))
    return max(0.0, 1.0 - min(abs(ratio) / 2.4, 1.0))


def _area_score(a: list[int], b: list[int]) -> float:
    area_a = max(_rect_area(a), 1)
    area_b = max(_rect_area(b), 1)
    ratio = math.log(area_a / area_b)
    return max(0.0, 1.0 - min(abs(ratio) / 3.0, 1.0))


def _match_score(
    *,
    center_norm: float,
    iou: float,
    ph_cover: float,
    element_cover: float,
    aspect_score: float,
    area_score: float,
    element_to_placeholder_ratio: float,
) -> float:
    center_score = 1.0 / (1.0 + center_norm)
    oversized_penalty = 0.0
    if element_to_placeholder_ratio > MATCH_OVERSIZED_RATIO:
        oversized_penalty = min((element_to_placeholder_ratio - MATCH_OVERSIZED_RATIO) / 3.0, 0.35)
    score = (
        center_score * 0.35
        + iou * 0.2
        + ph_cover * 0.16
        + element_cover * 0.14
        + aspect_score * 0.08
        + area_score * 0.07
        - oversized_penalty
    )
    return max(0.0, min(score, 1.0))


def _prompt_for_element(element: dict[str, Any], image_area: int) -> str:
    bbox = element["bbox"]
    area_ratio = _rect_area(bbox) / max(image_area, 1)
    element_type = str(element.get("type", "image") or "image").strip().lower()
    content = str(element.get("content", "") or "").strip()
    if element_type == "table":
        return "table screenshot"
    if element_type == "equation":
        return content[:80] or "equation graphic"
    if content:
        return content[:80]
    if area_ratio <= 0.015:
        return "icon graphic"
    if area_ratio >= 0.18:
        return "screenshot image"
    return "graphic element"


def _asset_kind_for_element(element: dict[str, Any], image_area: int) -> str:
    bbox = element["bbox"]
    area_ratio = _rect_area(bbox) / max(image_area, 1)
    aspect = max(bbox[2] - bbox[0], 1) / max(bbox[3] - bbox[1], 1)
    element_type = str(element.get("type", "image") or "image").strip().lower()
    if element_type == "table":
        return "screenshot"
    if element_type == "equation":
        return "graphic"
    if area_ratio <= 0.015:
        return "icon"
    if 2.4 <= aspect <= 8.0 and area_ratio <= 0.06:
        return "logo"
    if area_ratio >= 0.18:
        return "screenshot"
    if area_ratio >= 0.08:
        return "illustration"
    return "graphic"


def _match_metrics(
    placeholder: dict[str, Any],
    element: dict[str, Any],
    image_diagonal: float,
) -> dict[str, Any]:
    ph_bbox = placeholder["bbox_px"]
    element_bbox = element["bbox"]
    ph_area = max(_rect_area(ph_bbox), 1)
    element_area = max(_rect_area(element_bbox), 1)
    inter = _intersection_area(ph_bbox, element_bbox)
    iou = _iou(ph_bbox, element_bbox)
    ph_cover = inter / ph_area
    element_cover = inter / element_area
    center_norm = _center_distance_norm(ph_bbox, element_bbox, image_diagonal)
    aspect_score = _aspect_score(ph_bbox, element_bbox)
    area_score = _area_score(ph_bbox, element_bbox)
    element_to_placeholder_ratio = element_area / max(ph_area, 1)
    score = _match_score(
        center_norm=center_norm,
        iou=iou,
        ph_cover=ph_cover,
        element_cover=element_cover,
        aspect_score=aspect_score,
        area_score=area_score,
        element_to_placeholder_ratio=element_to_placeholder_ratio,
    )
    return {
        "placeholder_id": placeholder["placeholder_id"],
        "element_id": element["element_id"],
        "score": round(score, 6),
        "iou": round(iou, 6),
        "ph_cover": round(ph_cover, 6),
        "element_cover": round(element_cover, 6),
        "center_distance_norm": round(center_norm, 6),
        "placeholder_area": ph_area,
        "element_area": element_area,
        "placeholder_to_element_ratio": round(ph_area / max(element_area, 1), 6),
        "element_to_placeholder_ratio": round(element_to_placeholder_ratio, 6),
        "aspect_score": round(aspect_score, 6),
        "area_score": round(area_score, 6),
    }


def _is_oversized_match(metrics: dict[str, Any]) -> bool:
    element_ratio = float(metrics["element_to_placeholder_ratio"])
    iou = float(metrics["iou"])
    element_cover = float(metrics["element_cover"])
    return (
        element_ratio >= MATCH_OVERSIZED_RATIO_HARD
        or (element_ratio >= MATCH_OVERSIZED_RATIO and iou < MATCH_LOW_IOU_FOR_OVERSIZED)
        or (element_ratio >= 1.45 and element_cover < MATCH_LOW_ELEMENT_COVER)
    )


def _can_salvage_oversized_match(metrics: dict[str, Any]) -> bool:
    if not _is_oversized_match(metrics):
        return True
    return (
        float(metrics["ph_cover"]) >= SALVAGE_MIN_PLACEHOLDER_COVER
        and float(metrics["center_distance_norm"]) <= SALVAGE_MAX_CENTER_DISTANCE
    )


def _is_match_acceptable(metrics: dict[str, Any]) -> bool:
    if float(metrics["aspect_score"]) < MATCH_MIN_ASPECT_SCORE:
        return False
    if float(metrics["area_score"]) < MATCH_MIN_AREA_SCORE:
        return False
    if float(metrics["iou"]) < MATCH_MIN_IOU and float(metrics["ph_cover"]) < SALVAGE_MIN_PLACEHOLDER_COVER:
        return False
    if float(metrics["ph_cover"]) < MATCH_MIN_PLACEHOLDER_COVER:
        return False
    if float(metrics["placeholder_to_element_ratio"]) > 10.0:
        return False
    if float(metrics["element_to_placeholder_ratio"]) > MATCH_MAX_ELEMENT_TO_PLACEHOLDER_RATIO:
        return False
    if not _can_salvage_oversized_match(metrics):
        return False
    return True


def _count_related_placeholders(element_bbox: list[int], placeholders: list[dict[str, Any]]) -> int:
    total = 0
    for placeholder in placeholders:
        ph_bbox = placeholder["bbox_px"]
        ph_area = max(_rect_area(ph_bbox), 1)
        if _intersection_area(element_bbox, ph_bbox) / ph_area >= 0.12:
            total += 1
    return total


def _should_refine_candidate(
    metrics: dict[str, Any],
    *,
    element_bbox: list[int],
    placeholders: list[dict[str, Any]],
    depth: int,
    max_depth: int,
) -> bool:
    if depth >= max_depth:
        return False
    if _count_related_placeholders(element_bbox, placeholders) > 1:
        return True
    if _is_oversized_match(metrics):
        return True
    return float(metrics["iou"]) < 0.32 and float(metrics["ph_cover"]) > 0.72


def _load_json_file(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _discover_visual_elements(
    extracted_dir: Path,
    width: int,
    height: int,
    prefix: str,
) -> list[dict[str, Any]]:
    model_candidates = sorted(extracted_dir.rglob("*_model.json"))
    content_candidates = sorted(extracted_dir.rglob("*_content_list.json"))
    elements: list[dict[str, Any]] = []

    if model_candidates:
        pages = _load_json_file(model_candidates[0])
        if isinstance(pages, list):
            for page_index, page_blocks in enumerate(pages):
                if isinstance(page_blocks, dict):
                    page_blocks = page_blocks.get("layout_dets", [])
                if not isinstance(page_blocks, list):
                    continue
                for block_index, block in enumerate(page_blocks):
                    if not isinstance(block, dict):
                        continue
                    element_type = str(block.get("type", "") or "").strip().lower()
                    if element_type not in VISUAL_ELEMENT_TYPES:
                        continue
                    bbox = _bbox_from_any(block.get("bbox"), width, height)
                    if not bbox:
                        raw_poly = block.get("poly")
                        if isinstance(raw_poly, list):
                            try:
                                bbox = _bbox_from_poly([float(item) for item in raw_poly], width, height)
                            except (TypeError, ValueError):
                                bbox = None
                    if not bbox:
                        continue
                    elements.append(
                        {
                            "element_id": f"{prefix}_p{page_index:02d}_{block_index:03d}",
                            "type": element_type,
                            "bbox": bbox,
                            "content": str(block.get("content", "") or "").strip(),
                            "score": float(block.get("score") or 0.0),
                            "raw": block,
                        }
                    )

    if elements:
        return elements

    if content_candidates:
        rows = _load_json_file(content_candidates[0])
        if isinstance(rows, list):
            for block_index, row in enumerate(rows):
                if not isinstance(row, dict):
                    continue
                element_type = str(row.get("type", "") or "").strip().lower()
                if element_type not in VISUAL_ELEMENT_TYPES:
                    continue
                bbox = _bbox_from_any(row.get("bbox"), width, height)
                if not bbox:
                    continue
                elements.append(
                    {
                        "element_id": f"{prefix}_c{block_index:03d}",
                        "type": element_type,
                        "bbox": bbox,
                        "content": str(row.get("text", "") or "").strip(),
                        "score": 0.0,
                        "raw": row,
                    }
                )
    return elements


class MineruClient:
    def __init__(
        self,
        *,
        base_url: str,
        api_key: str,
        model_version: str,
        language: str,
        enable_formula: bool,
        enable_table: bool,
        is_ocr: bool,
        poll_interval_seconds: float,
        timeout_seconds: int,
    ) -> None:
        self.base_url = (base_url or MINERU_DEFAULT_BASE_URL).rstrip("/")
        self.api_key = api_key.strip()
        self.model_version = (model_version or "vlm").strip() or "vlm"
        self.language = (language or "ch").strip() or "ch"
        self.enable_formula = bool(enable_formula)
        self.enable_table = bool(enable_table)
        self.is_ocr = bool(is_ocr)
        self.poll_interval_seconds = max(0.5, float(poll_interval_seconds or 2.0))
        self.timeout_seconds = max(30, int(timeout_seconds or 300))
        self.session = requests.Session()

    def _headers(self) -> dict[str, str]:
        return {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}",
        }

    def _parse_response(self, response: requests.Response, action: str) -> dict[str, Any]:
        try:
            payload = response.json()
        except ValueError as exc:
            response.raise_for_status()
            raise RuntimeError(f"MinerU {action} did not return JSON.") from exc
        if response.status_code != 200:
            raise RuntimeError(f"MinerU {action} failed with HTTP {response.status_code}: {payload}")
        code = payload.get("code", -1)
        try:
            code_value = int(code)
        except (TypeError, ValueError):
            code_value = -1
        if code_value != 0:
            raise RuntimeError(f"MinerU {action} failed: {payload.get('msg') or payload}")
        return payload

    def parse_local_file(
        self,
        *,
        file_path: Path,
        work_dir: Path,
        data_id: str,
    ) -> MineruParseResult:
        ensure_dir(work_dir)
        prepared_upload = _prepare_image_for_mineru_upload(file_path=file_path, work_dir=work_dir)
        request_payload = {
            "enable_formula": self.enable_formula,
            "language": self.language,
            "enable_table": self.enable_table,
            "files": [
                {
                    "name": prepared_upload.file_path.name,
                    "data_id": data_id,
                    "is_ocr": self.is_ocr,
                }
            ],
            "model_version": self.model_version,
        }
        response = self.session.post(
            f"{self.base_url}/file-urls/batch",
            headers=self._headers(),
            json=request_payload,
            timeout=min(self.timeout_seconds, 120),
        )
        payload = self._parse_response(response, "upload-url request")
        (work_dir / "request.json").write_text(json.dumps(request_payload, ensure_ascii=False, indent=2), encoding="utf-8")
        (work_dir / "request_response.json").write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        data = payload.get("data", {}) if isinstance(payload, dict) else {}
        batch_id = str(data.get("batch_id", "") or "").strip()
        file_urls = data.get("file_urls", [])
        upload_url = file_urls[0] if isinstance(file_urls, list) and file_urls else ""
        if not batch_id or not upload_url:
            raise RuntimeError(f"MinerU upload-url response is missing batch_id or file_urls: {payload}")

        with prepared_upload.file_path.open("rb") as handle:
            upload_response = self.session.put(upload_url, data=handle, timeout=min(self.timeout_seconds, 300))
        if upload_response.status_code not in {200, 201}:
            raise RuntimeError(
                f"MinerU upload failed with HTTP {upload_response.status_code}: {upload_response.text[:300]}"
            )

        result_payload = self._poll_batch_result(batch_id=batch_id, work_dir=work_dir)
        extract_result = ((result_payload.get("data") or {}).get("extract_result") or [])
        if not isinstance(extract_result, list) or not extract_result:
            raise RuntimeError(f"MinerU batch result does not contain extract_result: {result_payload}")
        result_row = extract_result[0]
        full_zip_url = str(result_row.get("full_zip_url", "") or "").strip()
        if not full_zip_url:
            raise RuntimeError(f"MinerU result does not contain full_zip_url: {result_payload}")

        zip_path = work_dir / "result.zip"
        download_response = self.session.get(full_zip_url, timeout=min(self.timeout_seconds, 300))
        download_response.raise_for_status()
        zip_path.write_bytes(download_response.content)

        extracted_dir = work_dir / "extracted"
        ensure_dir(extracted_dir)
        with zipfile.ZipFile(zip_path, "r") as archive:
            archive.extractall(extracted_dir)
        return MineruParseResult(
            extracted_dir=extracted_dir,
            prepared_upload=prepared_upload,
        )

    def _poll_batch_result(self, *, batch_id: str, work_dir: Path) -> dict[str, Any]:
        deadline = time.monotonic() + self.timeout_seconds
        last_payload: dict[str, Any] | None = None
        while time.monotonic() < deadline:
            response = self.session.get(
                f"{self.base_url}/extract-results/batch/{batch_id}",
                headers=self._headers(),
                timeout=min(self.timeout_seconds, 120),
            )
            payload = self._parse_response(response, "result polling")
            last_payload = payload
            (work_dir / "batch_status.json").write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

            rows = ((payload.get("data") or {}).get("extract_result") or [])
            if isinstance(rows, list) and rows:
                row = rows[0]
                state = str(row.get("state", "") or "").strip().lower()
                if state == "done":
                    return payload
                if state in {"failed", "error"}:
                    raise RuntimeError(f"MinerU parse failed: {row.get('err_msg') or row}")
            time.sleep(self.poll_interval_seconds)

        raise TimeoutError(f"MinerU parse timed out after {self.timeout_seconds} seconds. Last payload: {last_payload}")


def _crop_image(source_path: Path, crop_bbox: list[int], output_path: Path) -> list[int]:
    Image, _ = _load_pillow()
    with Image.open(source_path) as image:
        width, height = image.size
        margin = max(2, min(width, height) // 200)
        x1 = max(0, crop_bbox[0] - margin)
        y1 = max(0, crop_bbox[1] - margin)
        x2 = min(width, crop_bbox[2] + margin)
        y2 = min(height, crop_bbox[3] + margin)
        image.crop((x1, y1, x2, y2)).save(output_path)
    return [x1, y1, x2, y2]


def _crop_image_exact(source_path: Path, crop_bbox: list[int], output_path: Path) -> list[int]:
    Image, _ = _load_pillow()
    with Image.open(source_path) as image:
        width, height = image.size
        x1 = max(0, min(width - 1, int(crop_bbox[0])))
        y1 = max(0, min(height - 1, int(crop_bbox[1])))
        x2 = max(x1 + 1, min(width, int(crop_bbox[2])))
        y2 = max(y1 + 1, min(height, int(crop_bbox[3])))
        image.crop((x1, y1, x2, y2)).save(output_path)
    return [x1, y1, x2, y2]


def _expand_bbox(
    bbox: list[int],
    *,
    image_size: tuple[int, int],
    extra_ratio: float,
) -> list[int]:
    width, height = image_size
    bw = max(bbox[2] - bbox[0], 1)
    bh = max(bbox[3] - bbox[1], 1)
    cx = (bbox[0] + bbox[2]) / 2.0
    cy = (bbox[1] + bbox[3]) / 2.0
    expanded_w = bw * (1.0 + extra_ratio)
    expanded_h = bh * (1.0 + extra_ratio)
    x1 = max(0, int(round(cx - expanded_w / 2.0)))
    y1 = max(0, int(round(cy - expanded_h / 2.0)))
    x2 = min(width, int(round(cx + expanded_w / 2.0)))
    y2 = min(height, int(round(cy + expanded_h / 2.0)))
    if x2 <= x1:
        x2 = min(width, x1 + 1)
    if y2 <= y1:
        y2 = min(height, y1 + 1)
    return [x1, y1, x2, y2]


def _inset_bbox(
    bbox: list[int],
    *,
    image_size: tuple[int, int],
    inset_x_ratio: float,
    inset_top_ratio: float,
    inset_bottom_ratio: float,
) -> list[int]:
    width, height = image_size
    bw = max(bbox[2] - bbox[0], 1)
    bh = max(bbox[3] - bbox[1], 1)
    x1 = max(0, int(round(bbox[0] + bw * inset_x_ratio)))
    x2 = min(width, int(round(bbox[2] - bw * inset_x_ratio)))
    y1 = max(0, int(round(bbox[1] + bh * inset_top_ratio)))
    y2 = min(height, int(round(bbox[3] - bh * inset_bottom_ratio)))
    if x2 <= x1:
        x1, x2 = bbox[0], bbox[2]
    if y2 <= y1:
        y1, y2 = bbox[1], bbox[3]
    return [x1, y1, x2, y2]


def _union_bbox(boxes: list[list[int]]) -> Optional[list[int]]:
    if not boxes:
        return None
    x1 = min(box[0] for box in boxes)
    y1 = min(box[1] for box in boxes)
    x2 = max(box[2] for box in boxes)
    y2 = max(box[3] for box in boxes)
    if x2 <= x1 or y2 <= y1:
        return None
    return [x1, y1, x2, y2]


def _collect_foreground_components(
    *,
    source_rgba: Any,
    analysis_bbox: list[int],
    color_threshold: int = 18,
    max_edge: int = 256,
) -> list[dict[str, Any]]:
    Image, _ = _load_pillow()
    x1, y1, x2, y2 = analysis_bbox
    crop = source_rgba.crop((x1, y1, x2, y2)).convert("RGB")
    crop_w, crop_h = crop.size
    if crop_w <= 1 or crop_h <= 1:
        return []

    scale = min(1.0, max_edge / max(crop_w, crop_h))
    if scale < 1.0:
        small_size = (
            max(1, int(round(crop_w * scale))),
            max(1, int(round(crop_h * scale))),
        )
        crop = crop.resize(small_size, resample=_lanczos_resample(Image))
    small_w, small_h = crop.size
    pixels = crop.load()

    samples: list[tuple[int, int, int]] = []
    for sx in range(small_w):
        samples.append(pixels[sx, 0])
        samples.append(pixels[sx, small_h - 1])
    for sy in range(1, max(small_h - 1, 1)):
        samples.append(pixels[0, sy])
        samples.append(pixels[small_w - 1, sy])
    if not samples:
        return []
    bg = tuple(int(round(sum(channel) / len(samples))) for channel in zip(*samples))

    mask = [False] * (small_w * small_h)
    for sy in range(small_h):
        for sx in range(small_w):
            r, g, b = pixels[sx, sy]
            diff = abs(r - bg[0]) + abs(g - bg[1]) + abs(b - bg[2])
            chroma = max(r, g, b) - min(r, g, b)
            if diff >= color_threshold or chroma >= max(10, color_threshold // 2):
                mask[sy * small_w + sx] = True

    visited = [False] * len(mask)
    components: list[dict[str, Any]] = []
    min_area = max(12, int(small_w * small_h * 0.003))
    neighbors = ((1, 0), (-1, 0), (0, 1), (0, -1))

    for start_index, is_fg in enumerate(mask):
        if not is_fg or visited[start_index]:
            continue
        stack = [start_index]
        visited[start_index] = True
        area = 0
        min_x = max_x = start_index % small_w
        min_y = max_y = start_index // small_w
        while stack:
            current = stack.pop()
            cx = current % small_w
            cy = current // small_w
            area += 1
            min_x = min(min_x, cx)
            max_x = max(max_x, cx)
            min_y = min(min_y, cy)
            max_y = max(max_y, cy)
            for dx, dy in neighbors:
                nx = cx + dx
                ny = cy + dy
                if nx < 0 or ny < 0 or nx >= small_w or ny >= small_h:
                    continue
                next_index = ny * small_w + nx
                if visited[next_index] or not mask[next_index]:
                    continue
                visited[next_index] = True
                stack.append(next_index)
        if area < min_area:
            continue
        scale_x = crop_w / small_w
        scale_y = crop_h / small_h
        component_bbox = [
            x1 + int(round(min_x * scale_x)),
            y1 + int(round(min_y * scale_y)),
            x1 + int(round((max_x + 1) * scale_x)),
            y1 + int(round((max_y + 1) * scale_y)),
        ]
        components.append(
            {
                "bbox": component_bbox,
                "area": area,
                "touches_left": min_x <= 1,
                "touches_right": max_x >= small_w - 2,
                "touches_top": min_y <= 1,
                "touches_bottom": max_y >= small_h - 2,
            }
        )

    return components


def _component_score(component_bbox: list[int], analysis_bbox: list[int]) -> float:
    box_area = max(_rect_area(component_bbox), 1)
    acx = (analysis_bbox[0] + analysis_bbox[2]) / 2.0
    acy = (analysis_bbox[1] + analysis_bbox[3]) / 2.0
    bcx = (component_bbox[0] + component_bbox[2]) / 2.0
    bcy = (component_bbox[1] + component_bbox[3]) / 2.0
    diagonal = math.hypot(max(analysis_bbox[2] - analysis_bbox[0], 1), max(analysis_bbox[3] - analysis_bbox[1], 1))
    center_score = 1.0 - min(math.hypot(bcx - acx, bcy - acy) / max(diagonal, 1.0), 1.0)
    return box_area * (0.55 + center_score)


def _expand_bbox_from_base(
    bbox: list[int],
    *,
    base_bbox: list[int],
    image_size: tuple[int, int],
    expand_x_ratio: float,
    expand_top_ratio: float,
    expand_bottom_ratio: float,
) -> list[int]:
    width, height = image_size
    bw = max(base_bbox[2] - base_bbox[0], 1)
    bh = max(base_bbox[3] - base_bbox[1], 1)
    x1 = max(0, int(round(bbox[0] - bw * expand_x_ratio)))
    x2 = min(width, int(round(bbox[2] + bw * expand_x_ratio)))
    y1 = max(0, int(round(bbox[1] - bh * expand_top_ratio)))
    y2 = min(height, int(round(bbox[3] + bh * expand_bottom_ratio)))
    if x2 <= x1:
        x1, x2 = bbox[0], bbox[2]
    if y2 <= y1:
        y1, y2 = bbox[1], bbox[3]
    return [x1, y1, x2, y2]


def _component_guided_cutout_bbox(
    *,
    source_rgba: Any,
    placeholder_bbox: list[int],
    element_bbox: list[int],
    image_size: tuple[int, int],
    asset_kind: str,
) -> Optional[list[int]]:
    analysis_bbox = _intersect_bbox(
        element_bbox,
        _expand_bbox(placeholder_bbox, image_size=image_size, extra_ratio=0.18),
    )
    if not analysis_bbox or _rect_area(analysis_bbox) <= 16:
        return None

    components = _collect_foreground_components(
        source_rgba=source_rgba,
        analysis_bbox=analysis_bbox,
        color_threshold=16 if asset_kind == "screenshot" else 14,
    )
    if not components:
        return None

    components.sort(key=lambda item: _component_score(item["bbox"], analysis_bbox), reverse=True)

    if asset_kind == "screenshot":
        body_candidates = [
            item
            for item in components
            if not item["touches_left"] and not item["touches_right"] and not item["touches_bottom"]
        ]
        if not body_candidates:
            body_candidates = [item for item in components if not item["touches_bottom"]]
        if not body_candidates:
            body_candidates = components

        body_candidates.sort(key=lambda item: _component_score(item["bbox"], analysis_bbox), reverse=True)
        primary = body_candidates[0]["bbox"]
        safe_box = _inset_bbox(
            placeholder_bbox,
            image_size=image_size,
            inset_x_ratio=SCREENSHOT_SUBJECT_SAFE_X_RATIO,
            inset_top_ratio=SCREENSHOT_SUBJECT_SAFE_TOP_RATIO,
            inset_bottom_ratio=SCREENSHOT_SUBJECT_SAFE_BOTTOM_RATIO,
        )
        safe_cx = (safe_box[0] + safe_box[2]) / 2.0
        primary_cx = (primary[0] + primary[2]) / 2.0
        placeholder_width = max(placeholder_bbox[2] - placeholder_bbox[0], 1)
        center_bias_cap = placeholder_width * 0.03
        subject_center_x = safe_cx + max(-center_bias_cap, min(primary_cx - safe_cx, center_bias_cap))
        left_room = max(subject_center_x - safe_box[0], 1.0)
        right_room = max(safe_box[2] - subject_center_x, 1.0)
        primary_half_width = max(primary_cx - primary[0], primary[2] - primary_cx)
        max_symmetric_half_width = min(left_room, right_room) * 0.96
        padded_half_width = primary_half_width * (1.0 + SCREENSHOT_SUBJECT_PAD_X_RATIO)
        final_half_width = min(max_symmetric_half_width, max(primary_half_width, padded_half_width))

        if final_half_width >= primary_half_width:
            symmetric_bbox = [
                int(round(subject_center_x - final_half_width)),
                safe_box[1],
                int(round(subject_center_x + final_half_width)),
                safe_box[3],
            ]
            clipped = _intersect_bbox(symmetric_bbox, element_bbox)
            if clipped and _rect_area(clipped) >= max(int(_rect_area(placeholder_bbox) * 0.18), 16):
                return clipped
        return None

    selected_boxes: list[list[int]] = []
    analysis_cx = (analysis_bbox[0] + analysis_bbox[2]) / 2.0
    analysis_cy = (analysis_bbox[1] + analysis_bbox[3]) / 2.0
    analysis_diag = math.hypot(max(analysis_bbox[2] - analysis_bbox[0], 1), max(analysis_bbox[3] - analysis_bbox[1], 1))
    for item in components:
        box = item["bbox"]
        bcx = (box[0] + box[2]) / 2.0
        bcy = (box[1] + box[3]) / 2.0
        center_dist = math.hypot(bcx - analysis_cx, bcy - analysis_cy) / max(analysis_diag, 1.0)
        if center_dist > 0.36 and _rect_area(box) < _rect_area(analysis_bbox) * 0.06:
            continue
        selected_boxes.append(box)
        if len(selected_boxes) >= (1 if asset_kind == "screenshot" else 3):
            break

    union_bbox = _union_bbox(selected_boxes)
    if not union_bbox:
        return None

    expanded_union = _expand_bbox_from_base(
        union_bbox,
        base_bbox=placeholder_bbox,
        image_size=image_size,
        expand_x_ratio=0.03 if asset_kind == "screenshot" else 0.05,
        expand_top_ratio=0.02,
        expand_bottom_ratio=0.03 if asset_kind == "screenshot" else 0.05,
    )
    clipped_union = _intersect_bbox(expanded_union, analysis_bbox)
    if clipped_union and _rect_area(clipped_union) >= max(int(_rect_area(placeholder_bbox) * 0.2), 16):
        return clipped_union
    return None


def _compute_cutout_bbox(
    *,
    source_rgba: Any,
    placeholder_bbox: list[int],
    element: dict[str, Any],
    element_bbox: list[int],
    image_size: tuple[int, int],
    metrics: dict[str, Any],
) -> list[int]:
    if not _is_oversized_match(metrics):
        return list(element_bbox)

    image_area = image_size[0] * image_size[1]
    asset_kind = _asset_kind_for_element(element, image_area)

    component_guided = _component_guided_cutout_bbox(
        source_rgba=source_rgba,
        placeholder_bbox=placeholder_bbox,
        element_bbox=element_bbox,
        image_size=image_size,
        asset_kind=asset_kind,
    )
    if component_guided:
        return component_guided

    if asset_kind == "screenshot":
        inset_placeholder_bbox = _inset_bbox(
            placeholder_bbox,
            image_size=image_size,
            inset_x_ratio=SCREENSHOT_CUTOUT_INSET_X_RATIO,
            inset_top_ratio=SCREENSHOT_CUTOUT_INSET_TOP_RATIO,
            inset_bottom_ratio=SCREENSHOT_CUTOUT_INSET_BOTTOM_RATIO,
        )
        inset_clipped = _intersect_bbox(element_bbox, inset_placeholder_bbox)
        if inset_clipped and _rect_area(inset_clipped) >= max(int(_rect_area(placeholder_bbox) * 0.24), 16):
            return inset_clipped
    else:
        bleed_placeholder_bbox = _expand_bbox(
            placeholder_bbox,
            image_size=image_size,
            extra_ratio=GRAPHIC_CUTOUT_BLEED_RATIO,
        )
        bleed_clipped = _intersect_bbox(element_bbox, bleed_placeholder_bbox)
        if bleed_clipped and _rect_area(bleed_clipped) >= max(int(_rect_area(placeholder_bbox) * 0.32), 16):
            return bleed_clipped

        inset_placeholder_bbox = _inset_bbox(
            placeholder_bbox,
            image_size=image_size,
            inset_x_ratio=PLACEHOLDER_CUTOUT_INSET_X_RATIO,
            inset_top_ratio=PLACEHOLDER_CUTOUT_INSET_TOP_RATIO,
            inset_bottom_ratio=PLACEHOLDER_CUTOUT_INSET_BOTTOM_RATIO,
        )
        inset_clipped = _intersect_bbox(element_bbox, inset_placeholder_bbox)
        if inset_clipped and _rect_area(inset_clipped) >= max(int(_rect_area(placeholder_bbox) * 0.28), 16):
            return inset_clipped

    exact_clipped = _intersect_bbox(element_bbox, placeholder_bbox)
    if exact_clipped and _rect_area(exact_clipped) >= max(int(_rect_area(placeholder_bbox) * 0.35), 16):
        return exact_clipped

    conservative_bbox = _expand_bbox(
        placeholder_bbox,
        image_size=image_size,
        extra_ratio=min(PLACEHOLDER_CUTOUT_EXPAND_RATIO, 0.04),
    )
    clipped = _intersect_bbox(element_bbox, conservative_bbox)
    if clipped and _rect_area(clipped) >= max(int(_rect_area(placeholder_bbox) * 0.2), 16):
        return clipped

    fallback_bbox = _expand_bbox(
        placeholder_bbox,
        image_size=image_size,
        extra_ratio=PLACEHOLDER_CUTOUT_FALLBACK_EXPAND_RATIO,
    )
    clipped_fallback = _intersect_bbox(element_bbox, fallback_bbox)
    if clipped_fallback and _rect_area(clipped_fallback) >= 16:
        return clipped_fallback

    return fallback_bbox


def _refine_element_pool(
    *,
    source_image_path: Path,
    source_image_size: tuple[int, int],
    element: dict[str, Any],
    client: MineruClient,
    cache: dict[str, list[dict[str, Any]]],
    refine_root: Path,
    depth: int,
) -> list[dict[str, Any]]:
    bbox = element["bbox"]
    cache_key = ",".join(str(item) for item in bbox)
    cached = cache.get(cache_key)
    if cached is not None:
        return cached

    refine_dir = refine_root / f"depth_{depth}_{bbox[0]}_{bbox[1]}_{bbox[2]}_{bbox[3]}"
    ensure_dir(refine_dir)
    crop_path = refine_dir / "crop.png"
    global_crop_bbox = _crop_image(source_image_path, bbox, crop_path)
    parse_result = client.parse_local_file(
        file_path=crop_path,
        work_dir=refine_dir / "mineru",
        data_id=f"refine-{depth}-{uuid.uuid4().hex[:8]}",
    )
    crop_width = max(global_crop_bbox[2] - global_crop_bbox[0], 1)
    crop_height = max(global_crop_bbox[3] - global_crop_bbox[1], 1)
    refined = _discover_visual_elements(
        extracted_dir=parse_result.extracted_dir,
        width=parse_result.prepared_upload.upload_size[0],
        height=parse_result.prepared_upload.upload_size[1],
        prefix=f"refine_d{depth}",
    )
    refined = _remap_elements_to_size(
        refined,
        from_size=parse_result.prepared_upload.upload_size,
        to_size=(crop_width, crop_height),
    )

    global_elements: list[dict[str, Any]] = []
    for index, item in enumerate(refined):
        child_bbox = item["bbox"]
        global_bbox = [
            child_bbox[0] + global_crop_bbox[0],
            child_bbox[1] + global_crop_bbox[1],
            child_bbox[2] + global_crop_bbox[0],
            child_bbox[3] + global_crop_bbox[1],
        ]
        if _rect_area(global_bbox) <= 16:
            continue
        if _rect_area(global_bbox) >= _rect_area(bbox) * 0.96:
            continue
        child = dict(item)
        child["bbox"] = global_bbox
        child["element_id"] = f"{element['element_id']}_r{depth}_{index:03d}"
        child["refined_from"] = element["element_id"]
        global_elements.append(child)

    cache[cache_key] = global_elements
    return global_elements


def _search_unmatched_placeholder_region(
    *,
    placeholder: dict[str, Any],
    source_image_path: Path,
    source_image_size: tuple[int, int],
    client: MineruClient,
    search_root: Path,
) -> list[dict[str, Any]]:
    expanded_bbox = _expand_bbox(
        placeholder["bbox_px"],
        image_size=source_image_size,
        extra_ratio=UNMATCHED_PH_REGION_EXPAND_RATIO,
    )
    search_dir = search_root / placeholder["placeholder_id"]
    ensure_dir(search_dir)
    crop_path = search_dir / "crop.png"
    global_crop_bbox = _crop_image_exact(source_image_path, expanded_bbox, crop_path)
    parse_result = client.parse_local_file(
        file_path=crop_path,
        work_dir=search_dir / "mineru",
        data_id=f"unmatched-{placeholder['placeholder_id']}-{uuid.uuid4().hex[:8]}",
    )
    crop_width = max(global_crop_bbox[2] - global_crop_bbox[0], 1)
    crop_height = max(global_crop_bbox[3] - global_crop_bbox[1], 1)
    local_elements = _discover_visual_elements(
        extracted_dir=parse_result.extracted_dir,
        width=parse_result.prepared_upload.upload_size[0],
        height=parse_result.prepared_upload.upload_size[1],
        prefix=f"unmatched_{placeholder['placeholder_id']}",
    )
    local_elements = _remap_elements_to_size(
        local_elements,
        from_size=parse_result.prepared_upload.upload_size,
        to_size=(crop_width, crop_height),
    )

    global_elements: list[dict[str, Any]] = []
    for index, item in enumerate(local_elements):
        child_bbox = item["bbox"]
        global_bbox = [
            child_bbox[0] + global_crop_bbox[0],
            child_bbox[1] + global_crop_bbox[1],
            child_bbox[2] + global_crop_bbox[0],
            child_bbox[3] + global_crop_bbox[1],
        ]
        if _rect_area(global_bbox) <= 16:
            continue
        child = dict(item)
        child["bbox"] = global_bbox
        child["element_id"] = f"{placeholder['placeholder_id']}_u_{index:03d}"
        child["unmatched_region_bbox"] = global_crop_bbox
        global_elements.append(child)
    return global_elements


def _select_match_for_placeholder(
    *,
    placeholder: dict[str, Any],
    element_pool: dict[str, dict[str, Any]],
    used_element_ids: set[str],
    placeholders: list[dict[str, Any]],
    source_image_path: Path,
    source_image_size: tuple[int, int],
    client: MineruClient,
    cache: dict[str, list[dict[str, Any]]],
    refine_root: Path,
    max_refine_depth: int,
) -> Optional[dict[str, Any]]:
    diagonal = math.hypot(source_image_size[0], source_image_size[1])

    for depth in range(0, max_refine_depth + 1):
        candidates: list[dict[str, Any]] = []
        for element in element_pool.values():
            if element.get("inactive"):
                continue
            if element["element_id"] in used_element_ids:
                continue
            metrics = _match_metrics(placeholder, element, diagonal)
            candidates.append(
                {
                    "element": element,
                    "metrics": metrics,
                }
            )
        candidates.sort(key=lambda item: item["metrics"]["score"], reverse=True)
        if not candidates:
            return None

        best = candidates[0]
        best_element = best["element"]
        best_metrics = best["metrics"]
        if _should_refine_candidate(
            best_metrics,
            element_bbox=best_element["bbox"],
            placeholders=placeholders,
            depth=depth,
            max_depth=max_refine_depth,
        ):
            refined_children = _refine_element_pool(
                source_image_path=source_image_path,
                source_image_size=source_image_size,
                element=best_element,
                client=client,
                cache=cache,
                refine_root=refine_root,
                depth=depth + 1,
            )
            if refined_children:
                best_element["inactive"] = True
                for child in refined_children:
                    element_pool[child["element_id"]] = child
                continue

        if not _is_match_acceptable(best_metrics):
            return None
        return {
            "placeholder": placeholder,
            "element": best_element,
            "metrics": best_metrics,
        }

    return None


def _fill_unmatched_placeholders(
    *,
    placeholders: list[dict[str, Any]],
    matched_placeholder_ids: set[str],
    source_image_path: Path,
    source_image_size: tuple[int, int],
    client: MineruClient,
    search_root: Path,
) -> list[dict[str, Any]]:
    diagonal = math.hypot(source_image_size[0], source_image_size[1])
    recovered_matches: list[dict[str, Any]] = []
    for placeholder in placeholders:
        if placeholder["placeholder_id"] in matched_placeholder_ids:
            continue
        try:
            region_elements = _search_unmatched_placeholder_region(
                placeholder=placeholder,
                source_image_path=source_image_path,
                source_image_size=source_image_size,
                client=client,
                search_root=search_root,
            )
        except Exception:
            continue
        ranked: list[dict[str, Any]] = []
        for element in region_elements:
            metrics = _match_metrics(placeholder, element, diagonal)
            if not _is_match_acceptable(metrics):
                continue
            ranked.append({"placeholder": placeholder, "element": element, "metrics": metrics})
        ranked.sort(key=lambda item: item["metrics"]["score"], reverse=True)
        if ranked:
            recovered_matches.append(ranked[0])
    return recovered_matches


def _render_matches_to_disk(
    *,
    image_path: Path,
    image_size: tuple[int, int],
    assets_dir: Path,
    matches: list[dict[str, Any]],
) -> Path:
    Image, ImageDraw = _load_pillow()
    ensure_dir(assets_dir)
    cutouts_dir = assets_dir / "cutouts"
    debug_dir = assets_dir / "debug"
    ensure_dir(cutouts_dir)
    ensure_dir(debug_dir)

    width, height = image_size
    image_area = width * height
    export_rows: list[dict[str, Any]] = []
    with Image.open(image_path).convert("RGBA") as source_rgba:
        overlay = source_rgba.copy()
        draw = ImageDraw.Draw(overlay, "RGBA")
        for index, match in enumerate(matches):
            placeholder = match["placeholder"]
            element = match["element"]
            metrics = match["metrics"]
            bbox = _compute_cutout_bbox(
                source_rgba=source_rgba,
                placeholder_bbox=placeholder["bbox_px"],
                element=element,
                element_bbox=element["bbox"],
                image_size=image_size,
                metrics=metrics,
            )
            x1, y1, x2, y2 = bbox
            cutout = source_rgba.crop((x1, y1, x2, y2))
            cutout_name = f"{index:03d}_{placeholder['placeholder_id']}_{element['type']}.png"
            cutout_path = cutouts_dir / cutout_name
            cutout.save(cutout_path)

            draw.rectangle((x1, y1, x2, y2), outline=(0, 170, 255, 255), width=3)
            ph_bbox = placeholder["bbox_px"]
            draw.rectangle((ph_bbox[0], ph_bbox[1], ph_bbox[2], ph_bbox[3]), outline=(255, 96, 64, 255), width=3)

            export_rows.append(
                {
                    "id": index,
                    "placeholder_id": placeholder["placeholder_id"],
                    "asset_kind": _asset_kind_for_element(element, image_area),
                    "prompt": _prompt_for_element(element, image_area),
                    "score": metrics["score"],
                    "bbox": bbox,
                    "element_bbox": element["bbox"],
                    "placeholder_bbox_px": placeholder["bbox_px"],
                    "placeholder_bbox_slide": placeholder["bbox_slide"],
                    "cutout_path": str(cutout_path),
                    "source_backend": "edit",
                    "source_element_id": element["element_id"],
                    "source_element_type": element["type"],
                    "match_metrics": metrics,
                }
            )

        overlay.save(debug_dir / "matches_overlay.png")

    assets_json = assets_dir / "assets.json"
    assets_json.write_text(json.dumps(export_rows, ensure_ascii=False, indent=2), encoding="utf-8")
    return assets_json


def resolve_mineru_assets_json(
    *,
    image_path: Path,
    placeholders: list[dict[str, Any]],
    assets_dir: Path,
    base_url: str,
    api_key: str,
    model_version: str,
    language: str,
    enable_formula: bool,
    enable_table: bool,
    is_ocr: bool,
    poll_interval_seconds: float,
    timeout_seconds: int,
    max_refine_depth: int,
    force_reextract_assets: bool,
) -> Path:
    assets_dir = assets_dir.resolve()
    ensure_dir(assets_dir)
    assets_json = assets_dir / "assets.json"
    if assets_json.exists() and not force_reextract_assets:
        return assets_json

    width, height = read_image_size(image_path)
    image_size = (width, height)
    (assets_dir / "placeholders.json").write_text(json.dumps(placeholders, ensure_ascii=False, indent=2), encoding="utf-8")

    client = MineruClient(
        base_url=base_url,
        api_key=api_key,
        model_version=model_version,
        language=language,
        enable_formula=enable_formula,
        enable_table=enable_table,
        is_ocr=is_ocr,
        poll_interval_seconds=poll_interval_seconds,
        timeout_seconds=timeout_seconds,
    )

    root_dir = assets_dir / "edit_root"
    parse_result = client.parse_local_file(
        file_path=image_path,
        work_dir=root_dir,
        data_id=f"slide-{uuid.uuid4().hex[:8]}",
    )
    root_elements = _discover_visual_elements(
        extracted_dir=parse_result.extracted_dir,
        width=parse_result.prepared_upload.upload_size[0],
        height=parse_result.prepared_upload.upload_size[1],
        prefix="root",
    )
    root_elements = _remap_elements_to_size(
        root_elements,
        from_size=parse_result.prepared_upload.upload_size,
        to_size=image_size,
    )
    (assets_dir / "root_elements.json").write_text(
        json.dumps(root_elements, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    element_pool = {item["element_id"]: dict(item) for item in root_elements}
    used_element_ids: set[str] = set()
    refinement_cache: dict[str, list[dict[str, Any]]] = {}
    matches: list[dict[str, Any]] = []

    sorted_placeholders = sorted(
        placeholders,
        key=lambda row: (_rect_area(row["bbox_px"]), row["bbox_px"][1], row["bbox_px"][0]),
    )
    for placeholder in sorted_placeholders:
        match = _select_match_for_placeholder(
            placeholder=placeholder,
            element_pool=element_pool,
            used_element_ids=used_element_ids,
            placeholders=placeholders,
            source_image_path=image_path,
            source_image_size=image_size,
            client=client,
            cache=refinement_cache,
            refine_root=assets_dir / "edit_refine",
            max_refine_depth=max_refine_depth,
        )
        if not match:
            continue
        used_element_ids.add(match["element"]["element_id"])
        matches.append(match)

    recovered_matches = _fill_unmatched_placeholders(
        placeholders=sorted_placeholders,
        matched_placeholder_ids={row["placeholder"]["placeholder_id"] for row in matches},
        source_image_path=image_path,
        source_image_size=image_size,
        client=client,
        search_root=assets_dir / "edit_unmatched",
    )
    if recovered_matches:
        matches.extend(recovered_matches)

    matches.sort(key=lambda row: row["placeholder"]["placeholder_id"])
    (assets_dir / "matched_assets.json").write_text(
        json.dumps(matches, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return _render_matches_to_disk(
        image_path=image_path,
        image_size=image_size,
        assets_dir=assets_dir,
        matches=matches,
    )

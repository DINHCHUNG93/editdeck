from __future__ import annotations

from dataclasses import dataclass


@dataclass
class AnalyzeResult:
    text_shapes: int = 0
    overflow_like: int = 0
    major_overlap_pairs: int = 0
    visual_overlap_pairs: int = 0


@dataclass
class FixTuning:
    single_width_units_limit: float
    single_text_len_limit: int
    single_overflow_ratio_max: float
    single_scale_floor: float
    overflow_ignore_ratio: float
    overflow_scale_trigger_ratio: float
    overflow_scale_floor: float
    min_font_pt: float


def tuning_from_mode(mode: str) -> FixTuning:
    mode = mode.lower().strip()
    if mode == "safe":
        return FixTuning(
            single_width_units_limit=22.0,
            single_text_len_limit=24,
            single_overflow_ratio_max=1.50,
            single_scale_floor=0.86,
            overflow_ignore_ratio=1.08,
            overflow_scale_trigger_ratio=1.12,
            overflow_scale_floor=0.83,
            min_font_pt=8.5,
        )
    if mode == "aggressive":
        return FixTuning(
            single_width_units_limit=30.0,
            single_text_len_limit=34,
            single_overflow_ratio_max=1.80,
            single_scale_floor=0.76,
            overflow_ignore_ratio=1.03,
            overflow_scale_trigger_ratio=1.05,
            overflow_scale_floor=0.72,
            min_font_pt=8.0,
        )
    return FixTuning(
        single_width_units_limit=26.0,
        single_text_len_limit=28,
        single_overflow_ratio_max=1.65,
        single_scale_floor=0.85,
        overflow_ignore_ratio=1.05,
        overflow_scale_trigger_ratio=1.08,
        overflow_scale_floor=0.82,
        min_font_pt=9.0,
    )

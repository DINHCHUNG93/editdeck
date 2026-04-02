import argparse
import json
import sys
from pathlib import Path
from typing import Callable, Optional


def _parse_slide_count(raw: str) -> Optional[int]:
    value = (raw or "").strip().lower()
    if value in {"", "auto", "none", "null"}:
        return None
    try:
        count = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("slide_count must be an integer or `auto`.") from exc
    if count < 1 or count > 20:
        raise argparse.ArgumentTypeError("slide_count must be between 1 and 20.")
    return count


def _parse_information_density(raw: Optional[str]) -> str:
    value = (raw or "medium").strip().lower()
    if not value:
        value = "medium"
    if value not in {"auto", "low", "medium", "high", "extra"}:
        raise argparse.ArgumentTypeError("information_density must be one of: auto, low, medium, high, extra.")
    return value


def _load_requirement(inline_requirement: Optional[str], requirement_file: Optional[str]) -> str:
    direct = (inline_requirement or "").strip()
    if direct:
        return direct
    if requirement_file:
        return Path(requirement_file).read_text(encoding="utf-8").strip()
    raise ValueError("Please provide a requirement or use --requirement-file.")


def _load_style_template(path: Optional[str]) -> tuple[Optional[bytes], Optional[str]]:
    if not path:
        return None, None
    template_path = Path(path)
    suffix = template_path.suffix.lower()
    mime_map = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
    }
    return template_path.read_bytes(), mime_map.get(suffix, "application/octet-stream")


def _load_source_files(paths: Optional[list[str]]) -> list[object]:
    if not paths:
        return []
    from app.source_ingest import SourceFileInput

    source_files: list[object] = []
    for raw_path in paths:
        path = Path(raw_path).expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(f"Source file not found: {path}")
        source_files.append(SourceFileInput(name=path.name, data=path.read_bytes()))
    return source_files


class ProgressPrinter:
    def __init__(self) -> None:
        self.last_line = ""

    def __call__(self, payload: dict[str, object]) -> None:
        progress = int(payload.get("progress", 0) or 0)
        step = str(payload.get("step", "running") or "running")
        message = str(payload.get("message", "") or "")
        current_slide = int(payload.get("current_slide", 0) or 0)
        total_slides = int(payload.get("total_slides", 0) or 0)

        line = f"[{progress:>3}%] {step}: {message}"
        if total_slides > 0:
            line += f" | slides {current_slide}/{total_slides}"
        if line == self.last_line:
            return
        self.last_line = line
        print(line)


def _scaled_progress_callback(
    callback: Callable[[dict[str, object]], None],
    start: int,
    end: int,
) -> Callable[[dict[str, object]], None]:
    span = max(end - start, 1)

    def wrapped(payload: dict[str, object]) -> None:
        raw_progress = int(payload.get("progress", 0) or 0)
        scaled = start + int((max(0, min(100, raw_progress)) / 100) * span)
        mapped = dict(payload)
        mapped["progress"] = min(end, scaled)
        callback(mapped)

    return wrapped


def _save_json(path: Optional[str], payload: dict) -> None:
    if not path:
        return
    save_path = Path(path)
    save_path.parent.mkdir(parents=True, exist_ok=True)
    save_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"JSON: {save_path.resolve()}")


def _add_shared_generation_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("requirement", nargs="?", help="PPT requirement text.")
    parser.add_argument("--requirement-file", help="Load PPT requirement from a UTF-8 text file.")
    parser.add_argument("--slide-count", default="auto", help="Slide count, or `auto`.")
    parser.add_argument(
        "--information-density",
        default="medium",
        choices=["auto", "low", "medium", "high", "extra"],
        help="Outline information density. auto=no explicit density constraint; low=1-3, medium=3-5, high=5-7, extra=7-10 key points per slide.",
    )
    parser.add_argument("--style-description", help="Style description text.")
    parser.add_argument("--style-template", help="Style reference image path.")
    parser.add_argument(
        "--source-file",
        dest="source_files",
        action="append",
        help="Optional source document path. Repeatable. Supports .txt/.md/.pdf/.docx.",
    )
    parser.add_argument(
        "--export-mode",
        choices=["images", "ppt", "both"],
        default="both",
        help="Output mode: images / ppt / both.",
    )
    parser.add_argument("--output-dir", help="Output root directory.")
    parser.add_argument("--config-file", help="Config file path. Defaults to config/app.yaml.")
    parser.add_argument("--base-url", help="Override text model base_url.")
    parser.add_argument("--image-api-url", help="Override image model base_url.")
    parser.add_argument("--text-api-key", help="Override text model api_key.")
    parser.add_argument("--image-api-key", help="Override image model api_key.")
    parser.add_argument("--text-model", help="Override text model name.")
    parser.add_argument("--image-model", help="Override image model name.")
    parser.add_argument("--save-json", help="Write full result JSON to this file.")


def _add_editable_runtime_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--editable-base-url", help="Override editable model base_url.")
    parser.add_argument("--editable-api-key", help="Override editable model api_key.")
    parser.add_argument("--editable-model", help="Override editable model name.")
    parser.add_argument("--editable-prompt-file", help="Editable PPT prompt file path.")
    parser.add_argument("--editable-browser-path", help="Chrome/Chromium executable path.")
    parser.add_argument("--editable-download-timeout-ms", type=int, help="Browser download timeout in ms.")
    parser.add_argument("--editable-max-tokens", type=int, help="Max output tokens for editable model.")
    parser.add_argument("--editable-max-attempts", type=int, help="Max attempts per slide.")
    parser.add_argument("--editable-sleep-seconds", type=float, help="Sleep seconds between attempts.")
    parser.add_argument("--assets-json", help="Use an existing assets.json file. Single-image mode only.")
    parser.add_argument("--assets-dir", help="Asset output directory. Multi-slide mode writes into slide_xx subdirs.")
    parser.add_argument("--mineru-base-url", help="Override MINERU_BASE_URL.")
    parser.add_argument("--mineru-api-key", help="Override MINERU_API_KEY.")
    parser.add_argument("--mineru-model-version", help="Override MINERU_MODEL_VERSION.")
    parser.add_argument("--mineru-language", help="Override MINERU_LANGUAGE.")
    parser.add_argument("--mineru-poll-interval-seconds", type=float, help="MinerU polling interval seconds.")
    parser.add_argument("--mineru-timeout-seconds", type=int, help="MinerU task timeout in seconds.")
    parser.add_argument("--mineru-max-refine-depth", type=int, help="Max recursive MinerU refinement depth.")
    parser.add_argument(
        "--mineru-disable-formula",
        action="store_true",
        help="Disable MinerU formula extraction in VLM mode.",
    )
    parser.add_argument(
        "--mineru-disable-table",
        action="store_true",
        help="Disable MinerU table extraction in VLM mode.",
    )
    parser.add_argument(
        "--mineru-disable-ocr",
        action="store_true",
        help="Disable MinerU OCR during element parsing.",
    )
    parser.add_argument(
        "-edit",
        "--edit",
        dest="asset_backend",
        action="store_const",
        const="edit",
        help="Use Edit asset matching for PH placeholders.",
    )
    parser.add_argument("--force-reextract-assets", action="store_true", help="Force re-extract PH assets.")
    parser.add_argument("--disable-asset-reuse", action="store_true", help="Do not reuse one asset for multiple PHs.")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="python -m app.cli", description="PPT generation and editable PPT CLI.")
    subparsers = parser.add_subparsers(dest="command")

    generate = subparsers.add_parser("generate", help="Generate PPT images / PPT from requirement.")
    _add_shared_generation_args(generate)
    generate.add_argument("--editable-ppt", action="store_true", help="Continue to build editable PPT after images.")
    _add_editable_runtime_args(generate)

    editable = subparsers.add_parser("editable", help="Build editable PPT from existing slide images.")
    editable.add_argument("--run-dir", help="Existing run directory containing slide_*.png.")
    editable.add_argument(
        "--image",
        dest="images",
        action="append",
        help="Input slide image path. Repeat for multiple slides.",
    )
    editable.add_argument("--output-dir", required=True, help="Editable PPT output directory.")
    editable.add_argument("--config-file", help="Config file path. Defaults to config/app.yaml.")
    editable.add_argument("--save-json", help="Write full result JSON to this file.")
    _add_editable_runtime_args(editable)
    return parser


def _build_editable_runtime_config(editable_pipeline, args) -> object:
    return editable_pipeline.build_runtime_config(
        base_url=args.editable_base_url,
        api_key=args.editable_api_key,
        model=args.editable_model,
        prompt_file=args.editable_prompt_file,
        chrome_path=args.editable_browser_path,
        download_timeout_ms=args.editable_download_timeout_ms,
        max_tokens=args.editable_max_tokens,
        max_attempts=args.editable_max_attempts,
        sleep_seconds=args.editable_sleep_seconds,
        assets_json=args.assets_json,
        assets_dir=args.assets_dir,
        asset_backend=getattr(args, "asset_backend", None),
        mineru_base_url=args.mineru_base_url,
        mineru_api_key=args.mineru_api_key,
        mineru_model_version=args.mineru_model_version,
        mineru_language=args.mineru_language,
        mineru_enable_formula=False if getattr(args, "mineru_disable_formula", False) else None,
        mineru_enable_table=False if getattr(args, "mineru_disable_table", False) else None,
        mineru_is_ocr=False if getattr(args, "mineru_disable_ocr", False) else None,
        mineru_poll_interval_seconds=args.mineru_poll_interval_seconds,
        mineru_timeout_seconds=args.mineru_timeout_seconds,
        mineru_max_refine_depth=args.mineru_max_refine_depth,
        force_reextract_assets=args.force_reextract_assets if hasattr(args, "force_reextract_assets") else None,
        disable_asset_reuse=args.disable_asset_reuse if hasattr(args, "disable_asset_reuse") else None,
    )


def _validate_editable_backend_args(args: argparse.Namespace) -> None:
    asset_backend = (getattr(args, "asset_backend", None) or "").strip().lower()
    if asset_backend == "mineru":
        args.asset_backend = "edit"
        return
    if asset_backend and asset_backend != "edit":
        raise ValueError("`editable` backend only supports `edit`.")


def _run_generate(args: argparse.Namespace) -> int:
    from app.editable_ppt import EditableDeckPipeline
    from app.pipeline import PPTImagePipeline
    from app.settings import load_settings
    from app.source_ingest import SourceDocumentProcessor

    requirement = _load_requirement(args.requirement, args.requirement_file)
    if args.style_description and args.style_template:
        raise ValueError("`--style-description` and `--style-template` are mutually exclusive.")
    if args.editable_ppt:
        _validate_editable_backend_args(args)

    settings = load_settings(args.config_file)
    if args.output_dir:
        settings = settings.model_copy(
            update={
                "app": settings.app.model_copy(update={"output_root": str(Path(args.output_dir))}),
            }
        )

    pipeline = PPTImagePipeline(settings=settings)
    editable_pipeline = EditableDeckPipeline(settings=settings)
    source_processor = SourceDocumentProcessor(settings=settings)
    runtime_cfg = pipeline.build_runtime_config(
        base_url=args.base_url,
        image_api_url=args.image_api_url,
        text_api_key=args.text_api_key,
        image_api_key=args.image_api_key,
        text_model=args.text_model,
        image_model=args.image_model,
    )
    source_runtime_cfg = source_processor.build_runtime_config(
        text_provider=runtime_cfg.text_provider,
        text_base_url=runtime_cfg.text_base_url,
        text_api_key=runtime_cfg.text_api_key,
        text_model=runtime_cfg.text_model,
        mineru_base_url=args.mineru_base_url,
        mineru_api_key=args.mineru_api_key,
        mineru_model_version=args.mineru_model_version,
        mineru_language=args.mineru_language,
        mineru_enable_formula=False if getattr(args, "mineru_disable_formula", False) else None,
        mineru_enable_table=False if getattr(args, "mineru_disable_table", False) else None,
        mineru_is_ocr=False if getattr(args, "mineru_disable_ocr", False) else None,
        mineru_poll_interval_seconds=args.mineru_poll_interval_seconds,
        mineru_timeout_seconds=args.mineru_timeout_seconds,
    )
    editable_runtime_cfg = _build_editable_runtime_config(editable_pipeline, args) if args.editable_ppt else None
    style_bytes, style_mime = _load_style_template(args.style_template)
    source_payloads = _load_source_files(args.source_files)
    prepared_requirement = source_processor.prepare_requirement(
        user_requirement=requirement,
        source_files=source_payloads,
        runtime_cfg=source_runtime_cfg,
    )
    progress = ProgressPrinter()

    generation_progress = progress
    editable_progress = progress
    if args.editable_ppt:
        generation_progress = _scaled_progress_callback(progress, 0, 60)
        editable_progress = _scaled_progress_callback(progress, 60, 100)

    result = pipeline.run(
        user_requirement=prepared_requirement.final_requirement,
        slide_count=_parse_slide_count(args.slide_count),
        style_description=args.style_description,
        style_template_bytes=style_bytes,
        style_template_mime=style_mime,
        runtime_cfg=runtime_cfg,
        export_mode=args.export_mode,
        information_density=_parse_information_density(args.information_density),
        progress_callback=generation_progress,
    )

    if args.editable_ppt:
        editable_result = editable_pipeline.run_from_images(
            slide_images=[Path(slide.image_path) for slide in result.slides],
            runtime_cfg=editable_runtime_cfg,
            output_dir=Path(result.output_dir) / "editable_deck",
            progress_callback=editable_progress,
        )
        result = result.model_copy(update={"editable_deck": editable_result})

    print("")
    print(f"Title: {result.deck_title}")
    print(f"Run ID: {result.run_id}")
    print(f"Output Dir: {result.output_dir}")
    if result.log_dir:
        print(f"Log Dir: {result.log_dir}")
    if result.trace_path:
        print(f"Trace Log: {result.trace_path}")
    if result.progress_log_path:
        print(f"Progress Log: {result.progress_log_path}")
    print(f"Slides: {len(result.slides)}")
    if result.pptx_path:
        print(f"PPTX: {result.pptx_path}")
    if result.editable_deck:
        print(f"Editable PPTX: {result.editable_deck.pptx_path}")
        print(f"Editable Remaining PH: {result.editable_deck.total_remaining_ph_count}")
    for slide in result.slides:
        print(f"Image {slide.page:02d}: {slide.image_path}")

    _save_json(args.save_json, result.model_dump())
    return 0


def _run_editable(args: argparse.Namespace) -> int:
    from app.editable_ppt import EditableDeckPipeline
    from app.settings import load_settings

    _validate_editable_backend_args(args)
    settings = load_settings(args.config_file)
    editable_pipeline = EditableDeckPipeline(settings=settings)
    editable_runtime_cfg = _build_editable_runtime_config(editable_pipeline, args)
    progress = ProgressPrinter()

    if args.run_dir:
        result = editable_pipeline.run_from_run_dir(
            run_dir=Path(args.run_dir),
            runtime_cfg=editable_runtime_cfg,
            output_dir=Path(args.output_dir),
            progress_callback=progress,
        )
    else:
        images = [Path(path) for path in (args.images or [])]
        if not images:
            raise ValueError("`editable` requires `--run-dir` or at least one `--image`.")
        result = editable_pipeline.run_from_images(
            slide_images=images,
            runtime_cfg=editable_runtime_cfg,
            output_dir=Path(args.output_dir),
            progress_callback=progress,
        )

    print("")
    print(f"Editable Run ID: {result.run_id}")
    print(f"Output Dir: {result.output_dir}")
    print(f"Editable PPTX: {result.pptx_path}")
    print(f"Remaining PH: {result.total_remaining_ph_count}")
    for slide in result.slides:
        print(
            f"Slide {slide.page:02d}: attempt={slide.selected_attempt}, "
            f"assets={slide.asset_count}, remaining_ph={slide.remaining_ph_count}"
        )

    _save_json(args.save_json, result.model_dump())
    return 0


def main(argv: Optional[list[str]] = None) -> int:
    raw_argv = list(argv or sys.argv[1:])
    if not raw_argv or raw_argv[0] not in {"generate", "editable"}:
        raw_argv = ["generate", *raw_argv]

    parser = build_parser()
    args = parser.parse_args(raw_argv)

    try:
        if args.command == "editable":
            return _run_editable(args)
        return _run_generate(args)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

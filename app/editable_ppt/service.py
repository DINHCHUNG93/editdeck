from __future__ import annotations

import concurrent.futures
import re
import time
import traceback
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from threading import Lock
from typing import Any, Callable, Optional, Sequence

from app.schemas import EditableDeckResult, EditableSlideResult
from app.settings import Settings

from .assets import (
    SLIDE_HEIGHT_IN,
    SLIDE_WIDTH_IN,
    build_asset_runtime_script,
    build_browser_asset_manifest,
    read_image_size,
    write_json,
)
from .browser import (
    count_ph_text_in_pptx,
    ensure_dir,
    execute_html_and_download_pptx,
    execute_html_and_download_pptx_with_runtime,
)
from .codegen import SlideCodegenError, call_model_for_slide_code, load_prompt_text
from .mineru_assets import resolve_mineru_assets_json

ProgressCallback = Optional[Callable[[dict[str, Any]], None]]
SUPPORTED_IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp"}


@dataclass
class EditableRuntimeConfig:
    provider: str
    base_url: str
    api_key: str
    model: str
    prompt_file: Optional[str]
    chrome_path: Optional[str]
    download_timeout_ms: int
    max_tokens: int
    max_attempts: int
    sleep_seconds: float
    assets_json: Optional[str]
    assets_dir: Optional[str]
    asset_backend: str
    mineru_base_url: str
    mineru_api_key: str
    mineru_model_version: str
    mineru_language: str
    mineru_enable_formula: bool
    mineru_enable_table: bool
    mineru_is_ocr: bool
    mineru_poll_interval_seconds: float
    mineru_timeout_seconds: int
    mineru_max_refine_depth: int
    force_reextract_assets: bool
    disable_asset_reuse: bool
    render_workers: int


class EditableDeckPipeline:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.output_root = Path(settings.output_root).resolve()

    def build_runtime_config(
        self,
        base_url: Optional[str] = None,
        api_key: Optional[str] = None,
        model: Optional[str] = None,
        prompt_file: Optional[str] = None,
        chrome_path: Optional[str] = None,
        download_timeout_ms: Optional[int] = None,
        max_tokens: Optional[int] = None,
        max_attempts: Optional[int] = None,
        sleep_seconds: Optional[float] = None,
        assets_json: Optional[str] = None,
        assets_dir: Optional[str] = None,
        asset_backend: Optional[str] = None,
        mineru_base_url: Optional[str] = None,
        mineru_api_key: Optional[str] = None,
        mineru_model_version: Optional[str] = None,
        mineru_language: Optional[str] = None,
        mineru_enable_formula: Optional[bool] = None,
        mineru_enable_table: Optional[bool] = None,
        mineru_is_ocr: Optional[bool] = None,
        mineru_poll_interval_seconds: Optional[float] = None,
        mineru_timeout_seconds: Optional[int] = None,
        mineru_max_refine_depth: Optional[int] = None,
        force_reextract_assets: Optional[bool] = None,
        disable_asset_reuse: Optional[bool] = None,
        render_workers: Optional[int] = None,
    ) -> EditableRuntimeConfig:
        resolved_backend = (asset_backend or self.settings.editable_ppt_asset_backend or "edit").strip().lower()
        if resolved_backend == "mineru":
            resolved_backend = "edit"

        cfg = EditableRuntimeConfig(
            provider=(self.settings.editable_ppt_provider or "openai").strip().lower(),
            base_url=(base_url or self.settings.resolved_editable_base_url).strip(),
            api_key=(api_key or self.settings.resolved_editable_api_key).strip(),
            model=(model or self.settings.editable_ppt_model).strip(),
            prompt_file=(prompt_file or self.settings.editable_ppt_prompt_file or "").strip() or None,
            chrome_path=(chrome_path or self.settings.editable_ppt_browser_path or "").strip() or None,
            download_timeout_ms=int(download_timeout_ms or self.settings.editable_ppt_download_timeout_ms),
            max_tokens=int(max_tokens or self.settings.editable_ppt_max_tokens),
            max_attempts=int(max_attempts or self.settings.editable_ppt_max_attempts),
            sleep_seconds=float(
                sleep_seconds if sleep_seconds is not None else self.settings.editable_ppt_sleep_seconds
            ),
            assets_json=(assets_json or "").strip() or None,
            assets_dir=(assets_dir or "").strip() or None,
            asset_backend=resolved_backend,
            mineru_base_url=(mineru_base_url or self.settings.resolved_mineru_base_url).strip(),
            mineru_api_key=(mineru_api_key or self.settings.resolved_mineru_api_key).strip(),
            mineru_model_version=(mineru_model_version or self.settings.mineru_model_version).strip(),
            mineru_language=(mineru_language or self.settings.mineru_language).strip(),
            mineru_enable_formula=bool(
                self.settings.mineru_enable_formula if mineru_enable_formula is None else mineru_enable_formula
            ),
            mineru_enable_table=bool(
                self.settings.mineru_enable_table if mineru_enable_table is None else mineru_enable_table
            ),
            mineru_is_ocr=bool(self.settings.mineru_is_ocr if mineru_is_ocr is None else mineru_is_ocr),
            mineru_poll_interval_seconds=float(
                self.settings.mineru_poll_interval_seconds
                if mineru_poll_interval_seconds is None
                else mineru_poll_interval_seconds
            ),
            mineru_timeout_seconds=int(
                self.settings.mineru_timeout_seconds if mineru_timeout_seconds is None else mineru_timeout_seconds
            ),
            mineru_max_refine_depth=int(
                self.settings.mineru_max_refine_depth if mineru_max_refine_depth is None else mineru_max_refine_depth
            ),
            force_reextract_assets=bool(False if force_reextract_assets is None else force_reextract_assets),
            disable_asset_reuse=bool(
                self.settings.editable_ppt_disable_asset_reuse
                if disable_asset_reuse is None
                else disable_asset_reuse
            ),
            render_workers=int(render_workers or 0),
        )
        if cfg.provider not in {"openai", "gemini"}:
            raise ValueError("Editable model provider must be `openai` or `gemini`.")
        if not cfg.base_url:
            raise ValueError("EDITABLE_PPT_BASE_URL cannot be empty.")
        if not cfg.api_key:
            raise ValueError("EDITABLE_PPT_API_KEY cannot be empty.")
        if not cfg.model:
            raise ValueError("EDITABLE_PPT_MODEL cannot be empty.")
        if cfg.max_attempts < 1:
            raise ValueError("editable max_attempts must be >= 1.")
        if cfg.download_timeout_ms < 1000:
            raise ValueError("editable download_timeout_ms must be >= 1000.")
        if cfg.render_workers < 0:
            raise ValueError("editable render_workers must be >= 0.")
        if cfg.asset_backend != "edit":
            raise ValueError("editable asset backend must be `edit`.")
        if cfg.mineru_poll_interval_seconds <= 0:
            raise ValueError("mineru_poll_interval_seconds must be > 0.")
        if cfg.mineru_timeout_seconds < 30:
            raise ValueError("mineru_timeout_seconds must be >= 30.")
        if cfg.mineru_max_refine_depth < 0:
            raise ValueError("mineru_max_refine_depth must be >= 0.")
        if not cfg.mineru_api_key:
            raise ValueError("Edit backend requires MINERU_API_KEY or --mineru-api-key.")
        return cfg

    def run_from_run_dir(
        self,
        run_dir: Path,
        runtime_cfg: EditableRuntimeConfig,
        output_dir: Optional[Path] = None,
        progress_callback: ProgressCallback = None,
    ) -> EditableDeckResult:
        slide_images = self.discover_slide_images(run_dir)
        target_dir = output_dir or (run_dir / "editable_deck")
        return self.run_from_images(
            slide_images=slide_images,
            runtime_cfg=runtime_cfg,
            output_dir=target_dir,
            progress_callback=progress_callback,
        )

    def run_from_images(
        self,
        slide_images: Sequence[Path],
        runtime_cfg: EditableRuntimeConfig,
        output_dir: Path,
        progress_callback: ProgressCallback = None,
    ) -> EditableDeckResult:
        images = [Path(path).resolve() for path in slide_images]
        if not images:
            raise ValueError("No slide images were provided for editable PPT generation.")
        for image in images:
            if not image.exists():
                raise FileNotFoundError(f"Slide image not found: {image}")

        output_root = output_dir.resolve()
        ensure_dir(output_root)
        prompt_text = load_prompt_text(Path(runtime_cfg.prompt_file) if runtime_cfg.prompt_file else None)
        run_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:8]
        assets_json_arg = Path(runtime_cfg.assets_json).resolve() if runtime_cfg.assets_json else None
        if assets_json_arg is not None and len(images) > 1:
            raise ValueError("--assets-json only supports single-image mode.")

        render_workers = runtime_cfg.render_workers or len(images)
        render_workers = max(1, min(render_workers, len(images)))

        def emit(
            step: str,
            message: str,
            progress: int,
            current_slide: int = 0,
            total_slides: int = 0,
            done: bool = False,
            error: str = "",
        ) -> None:
            if not progress_callback:
                return
            progress_callback(
                {
                    "step": step,
                    "message": message,
                    "progress": max(0, min(100, int(progress))),
                    "current_slide": current_slide,
                    "total_slides": total_slides,
                    "done": done,
                    "error": error,
                }
            )

        try:
            emit("editable_prepare", "Preparing editable PPT pipeline...", 3, 0, len(images))
            emit(
                "editable_prepare",
                (
                    "Using precomputed editable assets; editable rendering "
                    f"uses {render_workers} workers."
                    if assets_json_arg is not None
                    else f"Edit asset matching runs after placeholder capture; editable rendering uses {render_workers} workers."
                ),
                7,
                0,
                len(images),
            )

            builders_by_key: dict[str, str] = {}
            manifests_by_key: dict[str, list[dict[str, Any]]] = {}
            slide_results: list[EditableSlideResult] = []
            errors: list[str] = []
            render_completed = 0
            progress_lock = Lock()

            with concurrent.futures.ThreadPoolExecutor(max_workers=render_workers) as executor:
                futures: dict[concurrent.futures.Future[dict[str, Any]], int] = {}

                for index, image_path in enumerate(images, start=1):
                    emit(
                        "editable_assets",
                        (
                            f"Loading provided assets for slide {index}/{len(images)}..."
                            if assets_json_arg is not None
                            else f"Preparing Edit placeholder capture for slide {index}/{len(images)}..."
                        ),
                        self._asset_progress(index, len(images), 0),
                        index,
                        len(images),
                    )
                    try:
                        prepared = self._prepare_slide(
                            index=index,
                            image_path=image_path,
                            output_root=output_root,
                            runtime_cfg=runtime_cfg,
                            assets_json_arg=assets_json_arg,
                        )
                    except Exception as exc:  # noqa: BLE001
                        errors.append(f"slide_{index:02d}: {exc}")
                        continue

                    emit(
                        "editable_codegen",
                        f"Queued editable rendering for slide {index}/{len(images)}...",
                        self._asset_progress(index, len(images), 20),
                        index,
                        len(images),
                    )
                    future = executor.submit(
                        self._render_slide,
                        slide_index=index,
                        slide_key=prepared["slide_key"],
                        image_path=image_path,
                        slide_dir=prepared["slide_dir"],
                        prompt_text=prompt_text,
                        manifest=prepared["manifest"],
                        assets_json=prepared["assets_json"],
                        runtime_cfg=runtime_cfg,
                        total_slides=len(images),
                        emit=emit,
                    )
                    futures[future] = index

                for future in concurrent.futures.as_completed(futures):
                    index = futures[future]
                    try:
                        payload = future.result()
                        builders_by_key[payload["slide_key"]] = payload["builder_text"]
                        manifests_by_key[payload["slide_key"]] = payload["manifest"]
                        slide_results.append(payload["slide_result"])
                    except Exception as exc:  # noqa: BLE001
                        errors.append(f"slide_{index:02d}: {exc}")
                    finally:
                        with progress_lock:
                            render_completed += 1
                            emit(
                                "editable_render",
                                f"Editable rendering completed {render_completed}/{len(futures)} slides",
                                35 + int((render_completed / max(len(futures), 1)) * 53),
                                render_completed,
                                len(images),
                            )

            if errors:
                raise RuntimeError("; ".join(errors))

            slide_results.sort(key=lambda item: item.page)
            emit("editable_packaging", "Packaging full editable deck...", 90, len(images), len(images))
            combined_html_path = output_root / "editable_deck.html"
            combined_html_path.write_text(
                self._build_deck_html(
                    builders_by_key=builders_by_key,
                    manifests_by_key=manifests_by_key,
                    deck_file_name="editable_deck.pptx",
                    deck_title="Editable Deck",
                    allow_asset_reuse=not runtime_cfg.disable_asset_reuse,
                    drop_unmatched_placeholders=True,
                ),
                encoding="utf-8",
            )

            download_dir = output_root / "download"
            ensure_dir(download_dir)
            generated_path = execute_html_and_download_pptx(
                html_path=combined_html_path,
                download_dir=download_dir,
                chrome_path=runtime_cfg.chrome_path,
                timeout_ms=runtime_cfg.download_timeout_ms,
            )
            final_path = output_root / "editable_deck.pptx"
            generated_path.replace(final_path)
            total_remaining = count_ph_text_in_pptx(final_path)

            result = EditableDeckResult(
                run_id=run_id,
                output_dir=str(output_root),
                pptx_path=str(final_path),
                pptx_url=self._generated_url(final_path),
                total_remaining_ph_count=total_remaining,
                slides=slide_results,
            )
            write_json(output_root / "result.json", result.model_dump())
            emit("completed", "Editable PPT generation completed", 100, len(images), len(images), True)
            return result
        except Exception as exc:
            emit("failed", f"Editable PPT generation failed: {exc}", 100, done=True, error=str(exc))
            raise

    def discover_slide_images(self, run_dir: Path) -> list[Path]:
        base_dir = run_dir.resolve()
        if not base_dir.exists():
            raise FileNotFoundError(f"Run directory not found: {base_dir}")

        slide_images = [
            path
            for path in base_dir.iterdir()
            if path.is_file() and path.suffix.lower() in SUPPORTED_IMAGE_SUFFIXES and path.stem.startswith("slide_")
        ]
        slide_images.sort(key=self._sort_key)
        if not slide_images:
            raise ValueError(f"No slide images matching slide_*.png/jpg/webp were found in: {base_dir}")
        return slide_images

    def _prepare_slide(
        self,
        *,
        index: int,
        image_path: Path,
        output_root: Path,
        runtime_cfg: EditableRuntimeConfig,
        assets_json_arg: Optional[Path],
    ) -> dict[str, Any]:
        slide_key = f"slide_{index:02d}"
        slide_dir = output_root / slide_key
        ensure_dir(slide_dir)

        if assets_json_arg is None:
            return {
                "slide_key": slide_key,
                "slide_dir": slide_dir,
                "assets_json": None,
                "manifest": [],
            }

        assets_json = assets_json_arg.resolve()
        if not assets_json.exists():
            raise FileNotFoundError(f"Assets JSON not found: {assets_json}")
        manifest = build_browser_asset_manifest(assets_json, image_path)
        write_json(slide_dir / "browser_asset_manifest.json", manifest)
        return {
            "slide_key": slide_key,
            "slide_dir": slide_dir,
            "assets_json": assets_json,
            "manifest": manifest,
        }

    def _render_slide(
        self,
        *,
        slide_index: int,
        slide_key: str,
        image_path: Path,
        slide_dir: Path,
        prompt_text: str,
        manifest: list[dict[str, Any]],
        assets_json: Optional[Path],
        runtime_cfg: EditableRuntimeConfig,
        total_slides: int,
        emit: Callable[..., None],
    ) -> dict[str, Any]:
        attempt_result = self._run_best_attempt(
            slide_index=slide_index,
            slide_key=slide_key,
            image_path=image_path,
            slide_dir=slide_dir,
            prompt_text=prompt_text,
            manifest=manifest,
            runtime_cfg=runtime_cfg,
            total_slides=total_slides,
            emit=emit,
        )

        selected_builder_path = Path(str(attempt_result["selected_builder_path"]))
        builder_text = selected_builder_path.read_text(encoding="utf-8")
        preview_payload = attempt_result

        if assets_json is None:
            emit(
                "editable_assets",
                f"Generating Edit matched assets for slide {slide_index}/{total_slides}...",
                self._render_progress(slide_index, total_slides, 6),
                slide_index,
                total_slides,
            )
            placeholders = self._collect_placeholder_records(
                runtime_state=attempt_result.get("runtime_state"),
                slide_key=slide_key,
                image_path=image_path,
            )
            if placeholders:
                assets_json = resolve_mineru_assets_json(
                    image_path=image_path,
                    placeholders=placeholders,
                    assets_dir=self._resolve_assets_dir(runtime_cfg, slide_dir, slide_key, total_slides),
                    base_url=runtime_cfg.mineru_base_url,
                    api_key=runtime_cfg.mineru_api_key,
                    model_version=runtime_cfg.mineru_model_version,
                    language=runtime_cfg.mineru_language,
                    enable_formula=runtime_cfg.mineru_enable_formula,
                    enable_table=runtime_cfg.mineru_enable_table,
                    is_ocr=runtime_cfg.mineru_is_ocr,
                    poll_interval_seconds=runtime_cfg.mineru_poll_interval_seconds,
                    timeout_seconds=runtime_cfg.mineru_timeout_seconds,
                    max_refine_depth=runtime_cfg.mineru_max_refine_depth,
                    force_reextract_assets=runtime_cfg.force_reextract_assets,
                )
                manifest = build_browser_asset_manifest(assets_json, image_path)
                write_json(slide_dir / "browser_asset_manifest.json", manifest)
                preview_payload = self._render_preview_artifacts(
                    slide_key=slide_key,
                    builder_text=builder_text,
                    manifest=manifest,
                    preview_dir=slide_dir / "filled_preview",
                    deck_file_name=f"{slide_key}_filled_preview.pptx",
                    deck_title=f"{slide_key}_filled",
                    runtime_cfg=runtime_cfg,
                    capture_runtime_matches=False,
                )
            else:
                manifest = []
                write_json(slide_dir / "browser_asset_manifest.json", manifest)

        slide_result = EditableSlideResult(
            page=slide_index,
            image_path=str(image_path),
            assets_json_path=str(assets_json or ""),
            asset_count=len(manifest),
            selected_attempt=int(attempt_result["selected_attempt"]),
            attempt_dir=str(attempt_result["attempt_dir"]),
            builder_path=str(selected_builder_path),
            preview_html_path=str(preview_payload["preview_html_path"]),
            preview_pptx_path=str(preview_payload["preview_pptx_path"]),
            remaining_ph_count=int(preview_payload["remaining_ph_count"]),
        )
        return {
            "slide_key": slide_key,
            "manifest": manifest,
            "builder_text": builder_text,
            "slide_result": slide_result,
        }

    def _run_best_attempt(
        self,
        slide_index: int,
        slide_key: str,
        image_path: Path,
        slide_dir: Path,
        prompt_text: str,
        manifest: list[dict[str, Any]],
        runtime_cfg: EditableRuntimeConfig,
        total_slides: int,
        emit: Callable[..., None],
    ) -> dict[str, Any]:
        best_result: Optional[dict[str, Any]] = None
        last_error: Optional[Exception] = None
        retry_feedback: Optional[str] = None
        previous_builder: Optional[str] = None

        for attempt in range(1, runtime_cfg.max_attempts + 1):
            attempt_dir = slide_dir / f"attempt_{attempt:02d}"
            ensure_dir(attempt_dir)
            emit(
                "editable_render",
                f"Rendering preview for slide {slide_index}, attempt {attempt}...",
                self._render_progress(slide_index, total_slides, 0),
                slide_index,
                total_slides,
            )

            builder: Optional[str] = None
            raw_text: str = ""
            raw_response_path: Optional[Path] = None
            builder_path = attempt_dir / "build_slide.js"
            try:
                raw_text, builder = call_model_for_slide_code(
                    provider=runtime_cfg.provider,
                    base_url=runtime_cfg.base_url,
                    api_key=runtime_cfg.api_key,
                    image_path=image_path,
                    prompt_text=prompt_text,
                    model=runtime_cfg.model,
                    max_tokens=runtime_cfg.max_tokens,
                    retry_feedback=retry_feedback,
                    previous_builder=previous_builder,
                )
                raw_response_path = attempt_dir / "model_response.txt"
                raw_response_path.write_text(raw_text, encoding="utf-8")
                builder_path.write_text(builder, encoding="utf-8")

                preview_payload = self._render_preview_artifacts(
                    slide_key=slide_key,
                    builder_text=builder,
                    manifest=manifest,
                    preview_dir=attempt_dir,
                    deck_file_name=f"{slide_key}_preview.pptx",
                    deck_title=slide_key,
                    runtime_cfg=runtime_cfg,
                    capture_runtime_matches=True,
                )

                attempt_result = {
                    "selected_attempt": attempt,
                    "attempt_dir": str(attempt_dir),
                    "selected_builder_path": str(builder_path),
                    **preview_payload,
                }
                write_json(attempt_dir / "result.json", attempt_result)

                previous_builder = builder
                best_result = attempt_result
                break
            except SlideCodegenError as exc:
                raw_text = exc.raw_text
                if raw_text:
                    raw_response_path = attempt_dir / "model_response.txt"
                    raw_response_path.write_text(raw_text, encoding="utf-8")
                last_error = exc
                retry_feedback = self._build_retry_feedback(exc)
                (attempt_dir / "error.log").write_text(traceback.format_exc(), encoding="utf-8")
                result_payload: dict[str, Any] = {
                    "selected_attempt": attempt,
                    "error": f"{type(exc).__name__}: {exc}",
                    "retry_feedback": retry_feedback,
                }
                if raw_response_path is not None:
                    result_payload["model_response_path"] = str(raw_response_path)
                write_json(attempt_dir / "result.json", result_payload)
            except Exception as exc:  # noqa: BLE001
                last_error = exc
                if builder:
                    previous_builder = builder
                retry_feedback = self._build_retry_feedback(exc)
                (attempt_dir / "error.log").write_text(traceback.format_exc(), encoding="utf-8")
                result_payload = {
                    "selected_attempt": attempt,
                    "error": f"{type(exc).__name__}: {exc}",
                    "retry_feedback": retry_feedback,
                }
                if raw_response_path is not None:
                    result_payload["model_response_path"] = str(raw_response_path)
                if builder_path.exists():
                    result_payload["builder_path"] = str(builder_path)
                write_json(attempt_dir / "result.json", result_payload)
            if attempt < runtime_cfg.max_attempts and runtime_cfg.sleep_seconds > 0:
                time.sleep(runtime_cfg.sleep_seconds)

        if best_result is None:
            if last_error is not None:
                raise RuntimeError(
                    f"all {runtime_cfg.max_attempts} editable attempts failed for slide {slide_index}: {last_error}"
                ) from last_error
            raise RuntimeError(f"All editable attempts failed for slide: {image_path}")

        selected_builder_path = Path(str(best_result["selected_builder_path"]))
        selected_copy_path = slide_dir / "selected_build_slide.js"
        selected_copy_path.write_text(selected_builder_path.read_text(encoding="utf-8"), encoding="utf-8")
        best_result["selected_builder_path"] = str(selected_copy_path)
        write_json(slide_dir / "selected_result.json", best_result)
        return best_result

    def _render_preview_artifacts(
        self,
        *,
        slide_key: str,
        builder_text: str,
        manifest: list[dict[str, Any]],
        preview_dir: Path,
        deck_file_name: str,
        deck_title: str,
        runtime_cfg: EditableRuntimeConfig,
        capture_runtime_matches: bool,
    ) -> dict[str, Any]:
        ensure_dir(preview_dir)
        preview_html_path = preview_dir / "preview.html"
        preview_html_path.write_text(
            self._build_deck_html(
                builders_by_key={slide_key: builder_text},
                manifests_by_key={slide_key: manifest},
                deck_file_name=deck_file_name,
                deck_title=deck_title,
                allow_asset_reuse=not runtime_cfg.disable_asset_reuse,
                drop_unmatched_placeholders=True,
            ),
            encoding="utf-8",
        )
        preview_download_dir = preview_dir / "download"
        ensure_dir(preview_download_dir)
        runtime_state: dict[str, object] | None = None
        if capture_runtime_matches:
            preview_pptx_path, runtime_state = execute_html_and_download_pptx_with_runtime(
                html_path=preview_html_path,
                download_dir=preview_download_dir,
                chrome_path=runtime_cfg.chrome_path,
                timeout_ms=runtime_cfg.download_timeout_ms,
            )
        else:
            preview_pptx_path = execute_html_and_download_pptx(
                html_path=preview_html_path,
                download_dir=preview_download_dir,
                chrome_path=runtime_cfg.chrome_path,
                timeout_ms=runtime_cfg.download_timeout_ms,
            )

        payload = {
            "preview_html_path": str(preview_html_path),
            "preview_pptx_path": str(preview_pptx_path),
            "remaining_ph_count": count_ph_text_in_pptx(preview_pptx_path),
        }
        if runtime_state is not None:
            payload["runtime_state"] = runtime_state
        return payload

    def _collect_placeholder_records(
        self,
        *,
        runtime_state: Any,
        slide_key: str,
        image_path: Path,
    ) -> list[dict[str, Any]]:
        matches_by_slide = runtime_state.get("matches", {}) if isinstance(runtime_state, dict) else {}
        slide_matches = matches_by_slide.get(slide_key, [])
        if not isinstance(slide_matches, list):
            slide_matches = []
        if not slide_matches:
            return []

        image_width, image_height = read_image_size(image_path)
        placeholders: list[dict[str, Any]] = []
        for index, row in enumerate(slide_matches, start=1):
            if not isinstance(row, dict):
                continue
            placeholder = row.get("placeholder", {})
            if not isinstance(placeholder, dict):
                continue

            x = float(placeholder.get("x", 0.0) or 0.0)
            y = float(placeholder.get("y", 0.0) or 0.0)
            w = float(placeholder.get("w", 0.0) or 0.0)
            h = float(placeholder.get("h", 0.0) or 0.0)
            x = max(0.0, min(x, SLIDE_WIDTH_IN))
            y = max(0.0, min(y, SLIDE_HEIGHT_IN))
            w = max(0.001, min(w, max(SLIDE_WIDTH_IN - x, 0.001)))
            h = max(0.001, min(h, max(SLIDE_HEIGHT_IN - y, 0.001)))

            x1 = max(0, min(int(round(x * image_width / SLIDE_WIDTH_IN)), image_width - 1))
            y1 = max(0, min(int(round(y * image_height / SLIDE_HEIGHT_IN)), image_height - 1))
            x2 = max(x1 + 1, min(int(round((x + w) * image_width / SLIDE_WIDTH_IN)), image_width))
            y2 = max(y1 + 1, min(int(round((y + h) * image_height / SLIDE_HEIGHT_IN)), image_height))

            placeholders.append(
                {
                    "placeholder_id": f"ph_{index:03d}",
                    "bbox_slide": {
                        "x": round(x, 4),
                        "y": round(y, 4),
                        "w": round(w, 4),
                        "h": round(h, 4),
                    },
                    "bbox_px": [x1, y1, x2, y2],
                    "status": str(row.get("status", "") or ""),
                }
            )

        return placeholders

    def _resolve_assets_dir(
        self,
        runtime_cfg: EditableRuntimeConfig,
        slide_dir: Path,
        slide_key: str,
        total_slides: int,
    ) -> Path:
        if not runtime_cfg.assets_dir:
            return slide_dir / "edit_assets"
        root = Path(runtime_cfg.assets_dir).resolve()
        if total_slides == 1:
            return root
        return root / slide_key

    def _build_deck_html(
        self,
        builders_by_key: dict[str, str],
        manifests_by_key: dict[str, list[dict[str, Any]]],
        deck_file_name: str,
        deck_title: str,
        allow_asset_reuse: bool,
        drop_unmatched_placeholders: bool,
    ) -> str:
        renamed_functions: list[str] = []
        invocations: list[str] = []

        for slide_key in sorted(builders_by_key.keys(), key=self._sort_key):
            function_name = f"build_{slide_key}"
            renamed_functions.append(self._rename_builder_function(builders_by_key[slide_key], function_name))
            invocations.append(
                "\n".join(
                    [
                        "    {",
                        "        const slide = pptx.addSlide();",
                        "        patchSlideApi(slide);",
                        f"        slide.__slideKey = '{slide_key}';",
                        "        window.__slideRef = slide;",
                        f"        {function_name}(slide, pptx);",
                        "    }",
                    ]
                )
            )

        runtime_script = build_asset_runtime_script(
            manifests_by_key,
            allow_asset_reuse,
            drop_unmatched_placeholders,
        )
        functions_block = "\n\n".join(renamed_functions)
        invoke_block = "\n".join(invocations)

        return f"""<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>{deck_title}</title>
    <script src="https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js"></script>
  </head>
  <body>
    <button id="generate-btn" type="button" onclick="generateSlide()">Generate PPT</button>
    {runtime_script}
    <script>
function normalizeShapeType(shapeType) {{
    if (!shapeType) return 'rect';
    if (typeof shapeType !== 'string') return shapeType;
    const key = shapeType.replace(/[^a-z]/gi, '').toLowerCase();
    const aliasMap = {{
        circle: 'ellipse',
        oval: 'ellipse',
        rectangle: 'rect',
        roundedrect: 'roundRect',
        roundedrectangle: 'roundRect',
        straightline: 'line'
    }};
    return aliasMap[key] || shapeType;
}}

function normalizeShapeOptions(shapeType, options) {{
    const next = Object.assign({{}}, options || {{}});
    const num = (value, fallback = 0) => {{
        const parsed = Number(value);
        return Number.isFinite(parsed) ? parsed : fallback;
    }};
    next.x = num(next.x, 0);
    next.y = num(next.y, 0);
    next.w = num(next.w, 0);
    next.h = num(next.h, 0);
    if (next.w === 0 && next.h === 0) {{
        next.w = 0.001;
    }}
    if (shapeType !== 'line' && (!next.w || next.w < 0)) {{
        next.w = 0.001;
    }}
    if (shapeType !== 'line' && (!next.h || next.h < 0)) {{
        next.h = 0.001;
    }}
    return next;
}}

function patchSlideApi(slide) {{
    if (!slide || slide.__patchedAddShape) return;
    const originalAddShape = slide.addShape.bind(slide);
    slide.addShape = function patchedAddShape(shapeType, options) {{
        const resolvedType = normalizeShapeType(shapeType);
        const normalizedOptions = normalizeShapeOptions(resolvedType, options);
        try {{
            return originalAddShape(resolvedType, normalizedOptions);
        }} catch (error) {{
            return originalAddShape('rect', normalizeShapeOptions('rect', normalizedOptions));
        }}
    }};
    slide.__patchedAddShape = true;
}}

{functions_block}

async function generateSlide() {{
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';
    pptx.author = 'xinda-agent';
    pptx.subject = '{deck_title}';
    pptx.title = '{deck_title}';
{invoke_block}
    await pptx.writeFile({{ fileName: '{deck_file_name}' }});
}}
    </script>
  </body>
</html>
"""

    @staticmethod
    def _rename_builder_function(builder: str, new_name: str) -> str:
        return re.sub(r"function\s+buildSlide\b", f"function {new_name}", builder, count=1)

    @staticmethod
    def _sort_key(path_or_key: Any) -> tuple[int, str]:
        value = str(path_or_key)
        name = Path(value).stem if any(char in value for char in ("/", "\\")) else value
        match = re.search(r"slide_(\d+)", name)
        if match:
            return int(match.group(1)), name
        return 10**9, name

    def _generated_url(self, path: Path) -> str:
        try:
            relative_path = path.resolve().relative_to(self.output_root)
        except ValueError:
            return ""
        return "/generated/" + relative_path.as_posix()

    @staticmethod
    def _asset_progress(index: int, total: int, offset: int) -> int:
        base = 8 + int(((index - 1) / max(total, 1)) * 22)
        return min(34, base + min(offset, 10))

    @staticmethod
    def _render_progress(index: int, total: int, offset: int) -> int:
        base = 35 + int(((index - 1) / max(total, 1)) * 45)
        return min(89, base + min(offset, 12))

    @staticmethod
    def _build_retry_feedback(exc: Exception) -> str:
        return (
            "The previous preview failed at runtime. "
            "Fix the JavaScript and return a full fresh `buildSlide(slide, pptx)` function. "
            f"Runtime error: {type(exc).__name__}: {exc}"
        )

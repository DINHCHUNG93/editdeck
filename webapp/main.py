from datetime import datetime
from pathlib import Path
from threading import Lock, Thread
from typing import Any, Callable, Optional
from uuid import uuid4

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.editable_ppt import EditableDeckPipeline
from app.pipeline import PPTImagePipeline
from app.settings import get_settings
from app.source_ingest import SourceDocumentProcessor, SourceFileInput

settings = get_settings()
pipeline = PPTImagePipeline(settings=settings)
editable_pipeline = EditableDeckPipeline(settings=settings)
source_processor = SourceDocumentProcessor(settings=settings)

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"

app = FastAPI(
    title="One-Click PPT Generator",
    description="生成 PPT 图片、普通 PPT，并支持转换为整套可编辑 PPT。",
    version="0.5.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

Path(settings.output_root).mkdir(parents=True, exist_ok=True)
app.mount("/generated", StaticFiles(directory=settings.output_root), name="generated")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

JOBS: dict[str, dict[str, Any]] = {}
JOBS_LOCK = Lock()


def _now_iso() -> str:
    return datetime.utcnow().isoformat(timespec="seconds") + "Z"


def _parse_slide_count(slide_count_raw: Optional[str]) -> Optional[int]:
    raw = (slide_count_raw or "").strip().lower()
    if raw in {"", "auto", "none", "null"}:
        return None
    try:
        value = int(raw)
    except ValueError as exc:
        raise ValueError("slide_count 必须是整数或 auto。") from exc
    if value < 1 or value > 20:
        raise ValueError("slide_count 范围必须在 1-20。")
    return value


def _parse_information_density(information_density_raw: Optional[str]) -> str:
    raw = (information_density_raw or "medium").strip().lower()
    if not raw:
        raw = "medium"
    if raw not in {"auto", "low", "medium", "high", "extra"}:
        raise ValueError("information_density 必须是 auto、low、medium、high 或 extra。")
    return raw


def _validate_style_inputs(style_description: Optional[str], style_bytes: Optional[bytes]) -> None:
    if (style_description or "").strip() and style_bytes:
        raise ValueError("风格描述与风格模板图互斥，请二选一。")


def _parse_bool(value: Optional[str]) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "on"}


def _job_snapshot(job_id: str) -> dict[str, Any]:
    with JOBS_LOCK:
        if job_id not in JOBS:
            raise KeyError(job_id)
        return dict(JOBS[job_id])


def _update_job(job_id: str, **updates: Any) -> None:
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return
        job.update(updates)
        job["updated_at"] = _now_iso()


def _create_job(initial_message: str = "任务已创建，等待执行...") -> str:
    job_id = uuid4().hex
    now = _now_iso()
    with JOBS_LOCK:
        JOBS[job_id] = {
            "job_id": job_id,
            "state": "queued",
            "step": "queued",
            "message": initial_message,
            "progress": 0,
            "current_slide": 0,
            "total_slides": 0,
            "done": False,
            "error": "",
            "result": None,
            "created_at": now,
            "updated_at": now,
        }
    return job_id


def _scaled_progress_callback(
    target: Callable[[dict[str, Any]], None],
    start: int,
    end: int,
) -> Callable[[dict[str, Any]], None]:
    span = max(end - start, 1)

    def wrapped(payload: dict[str, Any]) -> None:
        raw_progress = int(payload.get("progress", 0) or 0)
        mapped = dict(payload)
        mapped["progress"] = min(end, start + int((max(0, min(100, raw_progress)) / 100) * span))
        target(mapped)

    return wrapped


def _resolve_run_dir(run_id: str) -> Path:
    run_dir = (Path(settings.output_root) / run_id).resolve()
    root_dir = Path(settings.output_root).resolve()
    try:
        run_dir.relative_to(root_dir)
    except ValueError as exc:
        raise ValueError("invalid run_id") from exc
    if not run_dir.exists():
        raise FileNotFoundError(f"run_id not found: {run_id}")
    return run_dir


def _build_editable_runtime_config(
    *,
    editable_base_url: Optional[str],
    editable_api_key: Optional[str],
    editable_model: Optional[str],
    editable_prompt_file: Optional[str],
    editable_browser_path: Optional[str],
    editable_download_timeout_ms: Optional[int],
    editable_max_tokens: Optional[int],
    editable_max_attempts: Optional[int],
    editable_sleep_seconds: Optional[float],
    assets_dir: Optional[str],
    asset_backend: Optional[str],
    mineru_base_url: Optional[str],
    mineru_api_key: Optional[str],
    mineru_model_version: Optional[str],
    mineru_language: Optional[str],
    mineru_enable_formula: Optional[bool],
    mineru_enable_table: Optional[bool],
    mineru_is_ocr: Optional[bool],
    mineru_poll_interval_seconds: Optional[float],
    mineru_timeout_seconds: Optional[int],
    mineru_max_refine_depth: Optional[int],
    force_reextract_assets: Optional[bool],
    disable_asset_reuse: Optional[bool],
):
    return editable_pipeline.build_runtime_config(
        base_url=editable_base_url,
        api_key=editable_api_key,
        model=editable_model,
        prompt_file=editable_prompt_file,
        chrome_path=editable_browser_path,
        download_timeout_ms=editable_download_timeout_ms,
        max_tokens=editable_max_tokens,
        max_attempts=editable_max_attempts,
        sleep_seconds=editable_sleep_seconds,
        assets_dir=assets_dir,
        asset_backend=asset_backend,
        mineru_base_url=mineru_base_url,
        mineru_api_key=mineru_api_key,
        mineru_model_version=mineru_model_version,
        mineru_language=mineru_language,
        mineru_enable_formula=mineru_enable_formula,
        mineru_enable_table=mineru_enable_table,
        mineru_is_ocr=mineru_is_ocr,
        mineru_poll_interval_seconds=mineru_poll_interval_seconds,
        mineru_timeout_seconds=mineru_timeout_seconds,
        mineru_max_refine_depth=mineru_max_refine_depth,
        force_reextract_assets=force_reextract_assets,
        disable_asset_reuse=disable_asset_reuse,
    )


def _build_source_runtime_config(
    *,
    text_provider: str,
    text_base_url: str,
    text_api_key: str,
    text_model: str,
    mineru_base_url: Optional[str],
    mineru_api_key: Optional[str],
    mineru_model_version: Optional[str],
    mineru_language: Optional[str],
    mineru_enable_formula: Optional[bool],
    mineru_enable_table: Optional[bool],
    mineru_is_ocr: Optional[bool],
    mineru_poll_interval_seconds: Optional[float],
    mineru_timeout_seconds: Optional[int],
):
    return source_processor.build_runtime_config(
        text_provider=text_provider,
        text_base_url=text_base_url,
        text_api_key=text_api_key,
        text_model=text_model,
        mineru_base_url=mineru_base_url,
        mineru_api_key=mineru_api_key,
        mineru_model_version=mineru_model_version,
        mineru_language=mineru_language,
        mineru_enable_formula=mineru_enable_formula,
        mineru_enable_table=mineru_enable_table,
        mineru_is_ocr=mineru_is_ocr,
        mineru_poll_interval_seconds=mineru_poll_interval_seconds,
        mineru_timeout_seconds=mineru_timeout_seconds,
    )


def _validate_editable_backend_args(
    *,
    asset_backend: Optional[str],
) -> None:
    backend = (asset_backend or "").strip().lower()
    if backend in {"", "edit", "mineru"}:
        return
    raise ValueError("editable backend only supports `edit`.")


def _run_generation_job(
    job_id: str,
    user_requirement: str,
    source_files: list[SourceFileInput],
    slide_count: Optional[int],
    information_density: str,
    style_description: Optional[str],
    style_bytes: Optional[bytes],
    style_mime: Optional[str],
    runtime_cfg,
    source_runtime_cfg,
    export_mode: str,
    generate_editable_ppt: bool,
    editable_runtime_cfg,
) -> None:
    _update_job(job_id, state="running", step="prepare", message="任务开始执行...", progress=1)

    def on_progress(payload: dict[str, Any]) -> None:
        _update_job(
            job_id,
            state="running",
            step=payload.get("step", "running"),
            message=payload.get("message", "处理中..."),
            progress=payload.get("progress", 0),
            current_slide=payload.get("current_slide", 0),
            total_slides=payload.get("total_slides", 0),
            done=payload.get("done", False),
            error=payload.get("error", ""),
        )

    try:
        prepared_requirement = source_processor.prepare_requirement(
            user_requirement=user_requirement,
            source_files=source_files,
            runtime_cfg=source_runtime_cfg,
        )
        generation_progress = on_progress
        editable_progress = on_progress
        if generate_editable_ppt:
            generation_progress = _scaled_progress_callback(on_progress, 0, 60)
            editable_progress = _scaled_progress_callback(on_progress, 60, 100)

        result = pipeline.run(
            user_requirement=prepared_requirement.final_requirement,
            slide_count=slide_count,
            style_description=style_description,
            style_template_bytes=style_bytes,
            style_template_mime=style_mime,
            runtime_cfg=runtime_cfg,
            export_mode=export_mode,
            information_density=information_density,
            progress_callback=generation_progress,
        )

        if generate_editable_ppt:
            editable_result = editable_pipeline.run_from_images(
                slide_images=[Path(slide.image_path) for slide in result.slides],
                runtime_cfg=editable_runtime_cfg,
                output_dir=Path(result.output_dir) / "editable_deck",
                progress_callback=editable_progress,
            )
            result = result.model_copy(update={"editable_deck": editable_result})

        _update_job(
            job_id,
            state="done",
            step="completed",
            message="生成完成",
            progress=100,
            done=True,
            error="",
            result=result.model_dump(),
        )
    except Exception as exc:
        _update_job(
            job_id,
            state="failed",
            step="failed",
            message=f"生成失败：{exc}",
            progress=100,
            done=True,
            error=str(exc),
            result=None,
        )


def _run_editable_job(job_id: str, run_id: str, editable_runtime_cfg) -> None:
    _update_job(job_id, state="running", step="editable_prepare", message="开始转换可编辑 PPT...", progress=1)

    def on_progress(payload: dict[str, Any]) -> None:
        _update_job(
            job_id,
            state="running",
            step=payload.get("step", "running"),
            message=payload.get("message", "处理中..."),
            progress=payload.get("progress", 0),
            current_slide=payload.get("current_slide", 0),
            total_slides=payload.get("total_slides", 0),
            done=payload.get("done", False),
            error=payload.get("error", ""),
        )

    try:
        run_dir = _resolve_run_dir(run_id)
        result = editable_pipeline.run_from_run_dir(
            run_dir=run_dir,
            runtime_cfg=editable_runtime_cfg,
            output_dir=run_dir / "editable_deck",
            progress_callback=on_progress,
        )
        _update_job(
            job_id,
            state="done",
            step="completed",
            message="可编辑 PPT 生成完成",
            progress=100,
            done=True,
            error="",
            result=result.model_dump(),
        )
    except Exception as exc:
        _update_job(
            job_id,
            state="failed",
            step="failed",
            message=f"可编辑 PPT 生成失败：{exc}",
            progress=100,
            done=True,
            error=str(exc),
            result=None,
        )


@app.get("/", include_in_schema=False)
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/generate")
async def generate_sync(
    user_requirement: str = Form(...),
    slide_count: Optional[str] = Form(default="auto"),
    information_density: Optional[str] = Form(default="medium"),
    style_description: Optional[str] = Form(default=None),
    style_template: Optional[UploadFile] = File(default=None),
    source_files: Optional[list[UploadFile]] = File(default=None),
    export_mode: str = Form(default="both"),
    generate_editable_ppt: Optional[str] = Form(default="false"),
    base_url: Optional[str] = Form(default=None),
    image_api_url: Optional[str] = Form(default=None),
    text_api_key: Optional[str] = Form(default=None),
    image_api_key: Optional[str] = Form(default=None),
    text_model: Optional[str] = Form(default=None),
    image_model: Optional[str] = Form(default=None),
    editable_base_url: Optional[str] = Form(default=None),
    editable_api_key: Optional[str] = Form(default=None),
    editable_model: Optional[str] = Form(default=None),
    editable_prompt_file: Optional[str] = Form(default=None),
    editable_browser_path: Optional[str] = Form(default=None),
    editable_download_timeout_ms: Optional[int] = Form(default=None),
    editable_max_tokens: Optional[int] = Form(default=None),
    editable_max_attempts: Optional[int] = Form(default=None),
    editable_sleep_seconds: Optional[float] = Form(default=None),
    assets_dir: Optional[str] = Form(default=None),
    asset_backend: Optional[str] = Form(default=None),
    mineru_base_url: Optional[str] = Form(default=None),
    mineru_api_key: Optional[str] = Form(default=None),
    mineru_model_version: Optional[str] = Form(default=None),
    mineru_language: Optional[str] = Form(default=None),
    mineru_enable_formula: Optional[str] = Form(default=None),
    mineru_enable_table: Optional[str] = Form(default=None),
    mineru_is_ocr: Optional[str] = Form(default=None),
    mineru_poll_interval_seconds: Optional[float] = Form(default=None),
    mineru_timeout_seconds: Optional[int] = Form(default=None),
    mineru_max_refine_depth: Optional[int] = Form(default=None),
    force_reextract_assets: Optional[str] = Form(default=None),
    disable_asset_reuse: Optional[str] = Form(default=None),
) -> dict[str, Any]:
    try:
        resolved_slide_count = _parse_slide_count(slide_count)
        resolved_information_density = _parse_information_density(information_density)
        runtime_cfg = pipeline.build_runtime_config(
            base_url=base_url,
            image_api_url=image_api_url,
            text_api_key=text_api_key,
            image_api_key=image_api_key,
            text_model=text_model,
            image_model=image_model,
        )
        source_runtime_cfg = _build_source_runtime_config(
            text_provider=runtime_cfg.text_provider,
            text_base_url=runtime_cfg.text_base_url,
            text_api_key=runtime_cfg.text_api_key,
            text_model=runtime_cfg.text_model,
            mineru_base_url=mineru_base_url,
            mineru_api_key=mineru_api_key,
            mineru_model_version=mineru_model_version,
            mineru_language=mineru_language,
            mineru_enable_formula=_parse_bool(mineru_enable_formula) if mineru_enable_formula is not None else None,
            mineru_enable_table=_parse_bool(mineru_enable_table) if mineru_enable_table is not None else None,
            mineru_is_ocr=_parse_bool(mineru_is_ocr) if mineru_is_ocr is not None else None,
            mineru_poll_interval_seconds=mineru_poll_interval_seconds,
            mineru_timeout_seconds=mineru_timeout_seconds,
        )
        editable_runtime_cfg = None
        if _parse_bool(generate_editable_ppt):
            _validate_editable_backend_args(
                asset_backend=asset_backend,
            )
            editable_runtime_cfg = _build_editable_runtime_config(
                editable_base_url=editable_base_url,
                editable_api_key=editable_api_key,
                editable_model=editable_model,
                editable_prompt_file=editable_prompt_file,
                editable_browser_path=editable_browser_path,
                editable_download_timeout_ms=editable_download_timeout_ms,
                editable_max_tokens=editable_max_tokens,
                editable_max_attempts=editable_max_attempts,
                editable_sleep_seconds=editable_sleep_seconds,
                assets_dir=assets_dir,
                asset_backend=asset_backend,
                mineru_base_url=mineru_base_url,
                mineru_api_key=mineru_api_key,
                mineru_model_version=mineru_model_version,
                mineru_language=mineru_language,
                mineru_enable_formula=_parse_bool(mineru_enable_formula) if mineru_enable_formula is not None else None,
                mineru_enable_table=_parse_bool(mineru_enable_table) if mineru_enable_table is not None else None,
                mineru_is_ocr=_parse_bool(mineru_is_ocr) if mineru_is_ocr is not None else None,
                mineru_poll_interval_seconds=mineru_poll_interval_seconds,
                mineru_timeout_seconds=mineru_timeout_seconds,
                mineru_max_refine_depth=mineru_max_refine_depth,
                force_reextract_assets=_parse_bool(force_reextract_assets) if force_reextract_assets is not None else None,
                disable_asset_reuse=_parse_bool(disable_asset_reuse) if disable_asset_reuse is not None else None,
            )
        style_bytes = await style_template.read() if style_template else None
        style_mime = style_template.content_type if style_template else None
        source_payloads = [
            SourceFileInput(name=upload.filename or "source", data=await upload.read())
            for upload in (source_files or [])
            if upload and (upload.filename or "").strip()
        ]
        _validate_style_inputs(style_description=style_description, style_bytes=style_bytes)
        prepared_requirement = source_processor.prepare_requirement(
            user_requirement=user_requirement,
            source_files=source_payloads,
            runtime_cfg=source_runtime_cfg,
        )

        result = pipeline.run(
            user_requirement=prepared_requirement.final_requirement,
            slide_count=resolved_slide_count,
            style_description=style_description,
            style_template_bytes=style_bytes,
            style_template_mime=style_mime,
            runtime_cfg=runtime_cfg,
            export_mode=export_mode,
            information_density=resolved_information_density,
            progress_callback=None,
        )
        if _parse_bool(generate_editable_ppt):
            editable_result = editable_pipeline.run_from_images(
                slide_images=[Path(slide.image_path) for slide in result.slides],
                runtime_cfg=editable_runtime_cfg,
                output_dir=Path(result.output_dir) / "editable_deck",
                progress_callback=None,
            )
            result = result.model_copy(update={"editable_deck": editable_result})
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"生成失败：{exc}") from exc
    return result.model_dump()


@app.post("/api/generate/start")
async def generate_start(
    user_requirement: str = Form(...),
    slide_count: Optional[str] = Form(default="auto"),
    information_density: Optional[str] = Form(default="medium"),
    style_description: Optional[str] = Form(default=None),
    style_template: Optional[UploadFile] = File(default=None),
    source_files: Optional[list[UploadFile]] = File(default=None),
    export_mode: str = Form(default="both"),
    generate_editable_ppt: Optional[str] = Form(default="false"),
    base_url: Optional[str] = Form(default=None),
    image_api_url: Optional[str] = Form(default=None),
    text_api_key: Optional[str] = Form(default=None),
    image_api_key: Optional[str] = Form(default=None),
    text_model: Optional[str] = Form(default=None),
    image_model: Optional[str] = Form(default=None),
    editable_base_url: Optional[str] = Form(default=None),
    editable_api_key: Optional[str] = Form(default=None),
    editable_model: Optional[str] = Form(default=None),
    editable_prompt_file: Optional[str] = Form(default=None),
    editable_browser_path: Optional[str] = Form(default=None),
    editable_download_timeout_ms: Optional[int] = Form(default=None),
    editable_max_tokens: Optional[int] = Form(default=None),
    editable_max_attempts: Optional[int] = Form(default=None),
    editable_sleep_seconds: Optional[float] = Form(default=None),
    assets_dir: Optional[str] = Form(default=None),
    asset_backend: Optional[str] = Form(default=None),
    mineru_base_url: Optional[str] = Form(default=None),
    mineru_api_key: Optional[str] = Form(default=None),
    mineru_model_version: Optional[str] = Form(default=None),
    mineru_language: Optional[str] = Form(default=None),
    mineru_enable_formula: Optional[str] = Form(default=None),
    mineru_enable_table: Optional[str] = Form(default=None),
    mineru_is_ocr: Optional[str] = Form(default=None),
    mineru_poll_interval_seconds: Optional[float] = Form(default=None),
    mineru_timeout_seconds: Optional[int] = Form(default=None),
    mineru_max_refine_depth: Optional[int] = Form(default=None),
    force_reextract_assets: Optional[str] = Form(default=None),
    disable_asset_reuse: Optional[str] = Form(default=None),
) -> dict[str, Any]:
    try:
        resolved_slide_count = _parse_slide_count(slide_count)
        resolved_information_density = _parse_information_density(information_density)
        runtime_cfg = pipeline.build_runtime_config(
            base_url=base_url,
            image_api_url=image_api_url,
            text_api_key=text_api_key,
            image_api_key=image_api_key,
            text_model=text_model,
            image_model=image_model,
        )
        source_runtime_cfg = _build_source_runtime_config(
            text_provider=runtime_cfg.text_provider,
            text_base_url=runtime_cfg.text_base_url,
            text_api_key=runtime_cfg.text_api_key,
            text_model=runtime_cfg.text_model,
            mineru_base_url=mineru_base_url,
            mineru_api_key=mineru_api_key,
            mineru_model_version=mineru_model_version,
            mineru_language=mineru_language,
            mineru_enable_formula=_parse_bool(mineru_enable_formula) if mineru_enable_formula is not None else None,
            mineru_enable_table=_parse_bool(mineru_enable_table) if mineru_enable_table is not None else None,
            mineru_is_ocr=_parse_bool(mineru_is_ocr) if mineru_is_ocr is not None else None,
            mineru_poll_interval_seconds=mineru_poll_interval_seconds,
            mineru_timeout_seconds=mineru_timeout_seconds,
        )
        editable_runtime_cfg = None
        if _parse_bool(generate_editable_ppt):
            _validate_editable_backend_args(
                asset_backend=asset_backend,
            )
            editable_runtime_cfg = _build_editable_runtime_config(
                editable_base_url=editable_base_url,
                editable_api_key=editable_api_key,
                editable_model=editable_model,
                editable_prompt_file=editable_prompt_file,
                editable_browser_path=editable_browser_path,
                editable_download_timeout_ms=editable_download_timeout_ms,
                editable_max_tokens=editable_max_tokens,
                editable_max_attempts=editable_max_attempts,
                editable_sleep_seconds=editable_sleep_seconds,
                assets_dir=assets_dir,
                asset_backend=asset_backend,
                mineru_base_url=mineru_base_url,
                mineru_api_key=mineru_api_key,
                mineru_model_version=mineru_model_version,
                mineru_language=mineru_language,
                mineru_enable_formula=_parse_bool(mineru_enable_formula) if mineru_enable_formula is not None else None,
                mineru_enable_table=_parse_bool(mineru_enable_table) if mineru_enable_table is not None else None,
                mineru_is_ocr=_parse_bool(mineru_is_ocr) if mineru_is_ocr is not None else None,
                mineru_poll_interval_seconds=mineru_poll_interval_seconds,
                mineru_timeout_seconds=mineru_timeout_seconds,
                mineru_max_refine_depth=mineru_max_refine_depth,
                force_reextract_assets=_parse_bool(force_reextract_assets) if force_reextract_assets is not None else None,
                disable_asset_reuse=_parse_bool(disable_asset_reuse) if disable_asset_reuse is not None else None,
            )
        style_bytes = await style_template.read() if style_template else None
        style_mime = style_template.content_type if style_template else None
        source_payloads = [
            SourceFileInput(name=upload.filename or "source", data=await upload.read())
            for upload in (source_files or [])
            if upload and (upload.filename or "").strip()
        ]
        _validate_style_inputs(style_description=style_description, style_bytes=style_bytes)

        job_id = _create_job()
        thread = Thread(
            target=_run_generation_job,
            args=(
                job_id,
                user_requirement,
                source_payloads,
                resolved_slide_count,
                resolved_information_density,
                style_description,
                style_bytes,
                style_mime,
                runtime_cfg,
                source_runtime_cfg,
                export_mode,
                _parse_bool(generate_editable_ppt),
                editable_runtime_cfg,
            ),
            daemon=True,
        )
        thread.start()
        return {"job_id": job_id, "state": "queued"}
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"任务启动失败：{exc}") from exc


@app.post("/api/editable/start")
async def editable_start(
    run_id: str = Form(...),
    editable_base_url: Optional[str] = Form(default=None),
    editable_api_key: Optional[str] = Form(default=None),
    editable_model: Optional[str] = Form(default=None),
    editable_prompt_file: Optional[str] = Form(default=None),
    editable_browser_path: Optional[str] = Form(default=None),
    editable_download_timeout_ms: Optional[int] = Form(default=None),
    editable_max_tokens: Optional[int] = Form(default=None),
    editable_max_attempts: Optional[int] = Form(default=None),
    editable_sleep_seconds: Optional[float] = Form(default=None),
    assets_dir: Optional[str] = Form(default=None),
    asset_backend: Optional[str] = Form(default=None),
    mineru_base_url: Optional[str] = Form(default=None),
    mineru_api_key: Optional[str] = Form(default=None),
    mineru_model_version: Optional[str] = Form(default=None),
    mineru_language: Optional[str] = Form(default=None),
    mineru_enable_formula: Optional[str] = Form(default=None),
    mineru_enable_table: Optional[str] = Form(default=None),
    mineru_is_ocr: Optional[str] = Form(default=None),
    mineru_poll_interval_seconds: Optional[float] = Form(default=None),
    mineru_timeout_seconds: Optional[int] = Form(default=None),
    mineru_max_refine_depth: Optional[int] = Form(default=None),
    force_reextract_assets: Optional[str] = Form(default=None),
    disable_asset_reuse: Optional[str] = Form(default=None),
) -> dict[str, Any]:
    try:
        _validate_editable_backend_args(
            asset_backend=asset_backend,
        )
        editable_runtime_cfg = _build_editable_runtime_config(
            editable_base_url=editable_base_url,
            editable_api_key=editable_api_key,
            editable_model=editable_model,
            editable_prompt_file=editable_prompt_file,
            editable_browser_path=editable_browser_path,
            editable_download_timeout_ms=editable_download_timeout_ms,
            editable_max_tokens=editable_max_tokens,
            editable_max_attempts=editable_max_attempts,
            editable_sleep_seconds=editable_sleep_seconds,
            assets_dir=assets_dir,
            asset_backend=asset_backend,
            mineru_base_url=mineru_base_url,
            mineru_api_key=mineru_api_key,
            mineru_model_version=mineru_model_version,
            mineru_language=mineru_language,
            mineru_enable_formula=_parse_bool(mineru_enable_formula) if mineru_enable_formula is not None else None,
            mineru_enable_table=_parse_bool(mineru_enable_table) if mineru_enable_table is not None else None,
            mineru_is_ocr=_parse_bool(mineru_is_ocr) if mineru_is_ocr is not None else None,
            mineru_poll_interval_seconds=mineru_poll_interval_seconds,
            mineru_timeout_seconds=mineru_timeout_seconds,
            mineru_max_refine_depth=mineru_max_refine_depth,
            force_reextract_assets=_parse_bool(force_reextract_assets) if force_reextract_assets is not None else None,
            disable_asset_reuse=_parse_bool(disable_asset_reuse) if disable_asset_reuse is not None else None,
        )
        _resolve_run_dir(run_id)
        job_id = _create_job(initial_message="可编辑 PPT 转换任务已创建，等待执行...")
        thread = Thread(
            target=_run_editable_job,
            args=(job_id, run_id, editable_runtime_cfg),
            daemon=True,
        )
        thread.start()
        return {"job_id": job_id, "state": "queued"}
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"任务启动失败：{exc}") from exc


@app.get("/api/generate/status/{job_id}")
def generate_status(job_id: str) -> dict[str, Any]:
    try:
        return _job_snapshot(job_id)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail="job not found") from exc

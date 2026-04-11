from __future__ import annotations

import os
import shutil
import zipfile
from pathlib import Path


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def resolve_browser_executable(chrome_path: str | None) -> str | None:
    candidates: list[str] = []

    if chrome_path:
        candidates.append(chrome_path)

    for env_name in (
        "EDITABLE_PPT_BROWSER_PATH",
        "CHROME_PATH",
        "GOOGLE_CHROME_BIN",
        "CHROMIUM_PATH",
        "BROWSER_PATH",
    ):
        env_value = os.getenv(env_name)
        if env_value:
            candidates.append(env_value)

    for name in (
        "google-chrome",
        "google-chrome-stable",
        "chromium",
        "chromium-browser",
        "chromium-headless-shell",
        "microsoft-edge",
        "microsoft-edge-stable",
        "brave-browser",
        "chrome",
    ):
        resolved = shutil.which(name)
        if resolved:
            candidates.append(resolved)

    seen: set[str] = set()
    for raw in candidates:
        value = str(raw or "").strip()
        if not value or value in seen:
            continue
        seen.add(value)

        path = Path(value).expanduser()
        if path.exists():
            return str(path)

        which_value = shutil.which(value)
        if which_value:
            return which_value

    return None


def _load_playwright():
    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
        from playwright.sync_api import sync_playwright
    except ImportError as exc:  # pragma: no cover - depends on runtime
        raise RuntimeError(
            "可编辑 PPT 功能需要 Playwright 运行时。请先安装 `playwright`，并准备可用的 Chromium/Chrome，或执行 `playwright install chromium`。"
        ) from exc
    return PlaywrightTimeoutError, sync_playwright


def wait_for_pptxgenjs(page, timeout_ms: int = 60000) -> None:
    polls = max(timeout_ms // 1000, 1)
    for _ in range(polls):
        loaded = page.evaluate("() => typeof window.PptxGenJS !== 'undefined'")
        if loaded:
            return
        page.wait_for_timeout(1000)
    raise TimeoutError("PptxGenJS did not load from CDN in time.")


def _read_runtime_state(page) -> dict[str, object]:
    return page.evaluate(
        """() => {
            const matches = window.__AUTO_ASSET_MATCHES || {};
            const usedSets = window.__AUTO_ASSET_USED_IDS || {};
            const usedIds = {};
            for (const [key, value] of Object.entries(usedSets)) {
                usedIds[key] = Array.isArray(value)
                    ? value
                    : (value && typeof value.values === 'function' ? Array.from(value.values()) : []);
            }
            return {
                matches,
                used_ids: usedIds,
            };
        }"""
    )


def _summarize_browser_errors(page_errors: list[str], console_errors: list[str]) -> str:
    seen: set[str] = set()
    merged: list[str] = []
    for source, rows in (("pageerror", page_errors), ("console", console_errors)):
        for row in rows:
            message = str(row or "").strip()
            if not message or message in seen:
                continue
            seen.add(message)
            merged.append(f"{source}: {message}")
            if len(merged) >= 3:
                return " | ".join(merged)
    return " | ".join(merged)


def execute_html_and_download_pptx_with_runtime(
    html_path: Path,
    download_dir: Path,
    chrome_path: str | None,
    timeout_ms: int,
) -> tuple[Path, dict[str, object]]:
    PlaywrightTimeoutError, sync_playwright = _load_playwright()
    download_dir.mkdir(parents=True, exist_ok=True)

    resolved_browser = resolve_browser_executable(chrome_path)
    playwright = None
    browser = None
    context = None
    page = None
    save_path: Path | None = None
    run_error: Exception | None = None
    runtime_state: dict[str, object] = {"matches": {}, "used_ids": {}}
    page_errors: list[str] = []
    console_errors: list[str] = []

    try:
        playwright = sync_playwright().start()
        launch_kwargs = {"headless": True}
        if resolved_browser:
            launch_kwargs["executable_path"] = resolved_browser
        browser = playwright.chromium.launch(**launch_kwargs)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.on("pageerror", lambda exc: page_errors.append(str(exc)))
        page.on(
            "console",
            lambda msg: console_errors.append(msg.text)
            if msg.type == "error"
            else None,
        )

        page.goto(html_path.resolve().as_uri(), wait_until="domcontentloaded", timeout=90000)
        wait_for_pptxgenjs(page, timeout_ms=60000)

        has_generate_fn = page.evaluate("() => typeof window.generateSlide === 'function'")
        if not has_generate_fn:
            error_summary = _summarize_browser_errors(page_errors, console_errors)
            detail = f" JavaScript runtime errors: {error_summary}" if error_summary else ""
            raise RuntimeError(f"Generated HTML does not define function generateSlide().{detail}")

        try:
            with page.expect_download(timeout=timeout_ms) as download_info:
                page.evaluate("() => window.generateSlide()")
            download = download_info.value
            runtime_state = _read_runtime_state(page)
        except PlaywrightTimeoutError as exc:
            error_summary = _summarize_browser_errors(page_errors, console_errors)
            detail = f" JavaScript runtime errors: {error_summary}" if error_summary else ""
            run_error = TimeoutError(f"No PPTX download event was detected after generateSlide().{detail}")
            run_error.__cause__ = exc
        except Exception as exc:  # noqa: BLE001
            error_summary = _summarize_browser_errors(page_errors, console_errors)
            detail = f" JavaScript runtime errors: {error_summary}" if error_summary else ""
            run_error = RuntimeError(f"generateSlide() execution failed: {exc}{detail}")

        if run_error is None:
            suggested_name = download.suggested_filename or "editable_deck.pptx"
            if not suggested_name.lower().endswith(".pptx"):
                suggested_name = f"{suggested_name}.pptx"

            save_path = download_dir / suggested_name
            download.save_as(str(save_path))
            if not save_path.exists() or save_path.stat().st_size == 0:
                run_error = RuntimeError("Downloaded PPTX is missing or empty.")

    except Exception as exc:  # pragma: no cover - depends on runtime
        message = str(exc)
        if "Executable doesn't exist" in message or "Please run the following command" in message:
            raise RuntimeError(
                "没有找到可用浏览器。请确认 `EDITABLE_PPT_BROWSER_PATH`、`CHROME_PATH`、`BROWSER_PATH` 或系统 PATH 已正确设置，或者执行 `playwright install chromium`。"
            ) from exc
        raise
    finally:
        if context is not None:
            try:
                context.close()
            except Exception:
                pass
        if browser is not None:
            try:
                browser.close()
            except Exception:
                pass
        if playwright is not None:
            try:
                playwright.stop()
            except Exception:
                pass

    if run_error is not None:
        raise run_error
    if save_path is None:
        raise RuntimeError("PPTX save path was not created.")
    return save_path, runtime_state


def execute_html_and_download_pptx(
    html_path: Path,
    download_dir: Path,
    chrome_path: str | None,
    timeout_ms: int,
) -> Path:
    save_path, _ = execute_html_and_download_pptx_with_runtime(
        html_path=html_path,
        download_dir=download_dir,
        chrome_path=chrome_path,
        timeout_ms=timeout_ms,
    )
    return save_path


def count_ph_text_in_pptx(pptx_path: Path) -> int:
    pattern = b">PH<"
    total = 0

    try:
        with zipfile.ZipFile(pptx_path, "r") as archive:
            slide_xmls = sorted(
                name
                for name in archive.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            )
            for name in slide_xmls:
                total += archive.read(name).count(pattern)
    except zipfile.BadZipFile as exc:
        raise RuntimeError(f"Invalid PPTX file: {pptx_path}") from exc

    return total

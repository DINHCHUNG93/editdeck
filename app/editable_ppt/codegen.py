from __future__ import annotations

import base64
import re
from pathlib import Path
from typing import Any, Optional

from app.model_api import chat_completion_text


PROMPT_PATH = Path(__file__).resolve().parent.parent / "prompts" / "editable_slide_builder.md"


class SlideCodegenError(ValueError):
    def __init__(self, message: str, *, raw_text: str = "") -> None:
        super().__init__(message)
        self.raw_text = raw_text


def load_prompt_text(prompt_file: Optional[Path]) -> str:
    path = prompt_file or PROMPT_PATH
    if not path.exists():
        raise FileNotFoundError(f"Editable PPT prompt file not found: {path}")
    return path.read_text(encoding="utf-8")


def encode_image(image_path: Path) -> str:
    return base64.b64encode(image_path.read_bytes()).decode("utf-8")


def image_mime_from_path(image_path: Path) -> str:
    ext = image_path.suffix.lower()
    if ext == ".png":
        return "image/png"
    if ext in {".jpg", ".jpeg"}:
        return "image/jpeg"
    if ext == ".webp":
        return "image/webp"
    return "application/octet-stream"


def normalize_content(content: Any) -> str:
    if isinstance(content, str):
        return content

    if isinstance(content, list):
        parts: list[str] = []
        for item in content:
            if isinstance(item, str):
                parts.append(item)
                continue
            if isinstance(item, dict):
                text = item.get("text")
                if isinstance(text, str):
                    parts.append(text)
                continue
            text = getattr(item, "text", None)
            if isinstance(text, str):
                parts.append(text)
        return "\n".join(parts)

    return str(content)


def _find_matching_brace(text: str, open_brace_index: int) -> int:
    depth = 0
    in_single = False
    in_double = False
    in_backtick = False
    in_line_comment = False
    in_block_comment = False
    escape = False

    index = open_brace_index
    while index < len(text):
        char = text[index]
        nxt = text[index + 1] if index + 1 < len(text) else ""

        if in_line_comment:
            if char == "\n":
                in_line_comment = False
            index += 1
            continue

        if in_block_comment:
            if char == "*" and nxt == "/":
                in_block_comment = False
                index += 2
                continue
            index += 1
            continue

        if in_single:
            if not escape and char == "'":
                in_single = False
            escape = (char == "\\") and not escape
            index += 1
            continue

        if in_double:
            if not escape and char == '"':
                in_double = False
            escape = (char == "\\") and not escape
            index += 1
            continue

        if in_backtick:
            if not escape and char == "`":
                in_backtick = False
            escape = (char == "\\") and not escape
            index += 1
            continue

        if char == "/" and nxt == "/":
            in_line_comment = True
            index += 2
            continue
        if char == "/" and nxt == "*":
            in_block_comment = True
            index += 2
            continue

        if char == "'":
            in_single = True
            escape = False
            index += 1
            continue
        if char == '"':
            in_double = True
            escape = False
            index += 1
            continue
        if char == "`":
            in_backtick = True
            escape = False
            index += 1
            continue

        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return index

        index += 1

    return -1


def _extract_fenced_block(text: str, language: str) -> Optional[str]:
    match = re.search(rf"```{language}\s*(.*?)```", text, re.IGNORECASE | re.DOTALL)
    if match:
        return match.group(1).strip()
    fallback = re.search(rf"```{language}\s*(.*)$", text, re.IGNORECASE | re.DOTALL)
    if fallback:
        return fallback.group(1).strip()
    return None


def _extract_function(text: str, name: str) -> Optional[str]:
    match = re.search(rf"(async\s+)?function\s+{name}\s*\([^)]*\)\s*\{{", text)
    if not match:
        return None
    open_brace = text.find("{", match.start())
    if open_brace < 0:
        return None
    close_brace = _find_matching_brace(text, open_brace)
    if close_brace < 0:
        return None
    return text[match.start() : close_brace + 1].strip()


def _has_function_marker(text: str, name: str) -> bool:
    return bool(re.search(rf"(async\s+)?function\s+{name}\s*\(", text))


def _extract_script_content(text: str) -> Optional[str]:
    matches = re.findall(r"<script[^>]*>(.*?)</script>", text, flags=re.IGNORECASE | re.DOTALL)
    if matches:
        return "\n\n".join(item.strip() for item in matches if item.strip())
    return None


def _strip_reasoning_blocks(text: str) -> str:
    cleaned = re.sub(r"<think>.*?</think>", "", text, flags=re.IGNORECASE | re.DOTALL)
    cleaned = re.sub(r"^\s*<think>.*$", "", cleaned, flags=re.IGNORECASE | re.MULTILINE)
    return cleaned.strip()


def _remove_function_definition(code: str, func_name: str) -> str:
    while True:
        match = re.search(rf"(async\s+)?function\s+{func_name}\s*\([^)]*\)\s*\{{", code)
        if not match:
            return code
        open_brace = code.find("{", match.start())
        if open_brace < 0:
            return code
        close_brace = _find_matching_brace(code, open_brace)
        if close_brace < 0:
            return code
        code = (code[: match.start()] + code[close_brace + 1 :]).strip()


def _indent(text: str, spaces: int) -> str:
    prefix = " " * spaces
    return "\n".join(prefix + line if line else "" for line in text.splitlines())


def _normalize_generate_slide(function_text: str) -> str:
    open_brace = function_text.find("{")
    close_brace = _find_matching_brace(function_text, open_brace)
    if open_brace < 0 or close_brace < 0:
        raise ValueError("Unable to parse generateSlide() function.")

    body = _remove_function_definition(function_text[open_brace + 1 : close_brace], "addPH")
    filtered_lines: list[str] = []
    for raw_line in body.replace("\r\n", "\n").splitlines():
        stripped = raw_line.strip()
        if not stripped:
            filtered_lines.append("")
            continue
        if "new PptxGenJS" in stripped:
            continue
        if "pptx.defineLayout" in stripped:
            continue
        if ".layout" in stripped and "pptx" in stripped:
            continue
        if "pptx.writeFile" in stripped or "pptx.write(" in stripped:
            continue
        if "document." in stripped or "addEventListener" in stripped:
            continue
        if re.search(r"^(const|let|var)\s+slide\s*=\s*.*addSlide", stripped):
            continue
        if re.search(r"^\s*slide\s*=\s*.*addSlide", stripped):
            continue
        filtered_lines.append(raw_line)

    normalized = "\n".join(filtered_lines).strip()
    if not normalized:
        raise ValueError("Model output did not contain usable slide-building code.")

    return (
        "function buildSlide(slide, pptx) {\n"
        "    window.__slideRef = slide;\n"
        f"{_indent(normalized, 4)}\n"
        "}"
    )


def _wrap_inline_slide_code(code: str) -> str:
    cleaned = _remove_function_definition(code, "addPH").strip()
    if not cleaned:
        raise ValueError("Model output did not contain usable slide-building code.")
    return (
        "function buildSlide(slide, pptx) {\n"
        "    window.__slideRef = slide;\n"
        f"{_indent(_sanitize_builder_code(cleaned), 4)}\n"
        "}"
    )


def _sanitize_builder_code(code: str) -> str:
    sanitized = code
    replacements = {
        r"\bpptx\.ShapeType\.circle\b": "pptx.ShapeType.ellipse",
        r"\bpptx\.ShapeType\.oval\b": "pptx.ShapeType.ellipse",
        r"\bpptx\.ShapeType\.roundedRect\b": "pptx.ShapeType.roundRect",
        r"\bpptx\.ShapeType\.roundedRectangle\b": "pptx.ShapeType.roundRect",
        r"\bpptx\.ShapeType\.rectangle\b": "pptx.ShapeType.rect",
        r"\bpptx\.ShapeType\.straightLine\b": "pptx.ShapeType.line",
    }
    for pattern, replacement in replacements.items():
        sanitized = re.sub(pattern, replacement, sanitized)
    return sanitized


def normalize_slide_builder(raw_text: str) -> str:
    cleaned_raw = _strip_reasoning_blocks(raw_text)
    candidates = [
        _extract_fenced_block(cleaned_raw, "javascript"),
        _extract_fenced_block(cleaned_raw, "js"),
        _extract_script_content(cleaned_raw),
        cleaned_raw.strip(),
    ]

    for candidate in candidates:
        if not candidate:
            continue

        build_slide = _extract_function(candidate, "buildSlide")
        if build_slide:
            return _sanitize_builder_code(_remove_function_definition(build_slide, "addPH").strip())
        if _has_function_marker(candidate, "buildSlide"):
            raise ValueError("Model output contained a buildSlide() declaration, but the function body was incomplete or truncated.")

        generate_slide = _extract_function(candidate, "generateSlide")
        if generate_slide:
            return _sanitize_builder_code(_normalize_generate_slide(generate_slide))
        if _has_function_marker(candidate, "generateSlide"):
            raise ValueError(
                "Model output contained a generateSlide() declaration, but the function body was incomplete or truncated."
            )

        if "slide.add" in candidate or "addPH(" in candidate:
            return _wrap_inline_slide_code(candidate)

    raise ValueError("Model output did not contain a usable buildSlide()/generateSlide() function.")


def call_model_for_slide_code(
    provider: str,
    base_url: str,
    api_key: str,
    image_path: Path,
    prompt_text: str,
    model: str,
    max_tokens: int,
    retry_feedback: Optional[str] = None,
    previous_builder: Optional[str] = None,
) -> tuple[str, str]:
    composed_prompt = prompt_text
    if retry_feedback:
        retry_block = [
            "## Retry Fix Context",
            "The previous `buildSlide(slide, pptx)` attempt failed or was not good enough.",
            "You must repair the code and return a complete fresh `buildSlide(slide, pptx)` function.",
            f"Failure details:\n{retry_feedback[:4000]}",
        ]
        if previous_builder:
            retry_block.append(
                "Previous code to fix:\n```javascript\n"
                + previous_builder[:12000]
                + "\n```"
            )
        composed_prompt = prompt_text + "\n\n" + "\n\n".join(retry_block)

    raw_text = chat_completion_text(
        provider=provider,
        base_url=base_url,
        api_key=api_key,
        model=model,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": composed_prompt},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:{image_mime_from_path(image_path)};base64,{encode_image(image_path)}"
                        },
                    },
                ],
            }
        ],
        temperature=None,
        max_tokens=max_tokens,
    )
    normalized_text = normalize_content(raw_text)
    try:
        builder = normalize_slide_builder(normalized_text)
    except ValueError as exc:
        raise SlideCodegenError(str(exc), raw_text=normalized_text) from exc
    return raw_text, builder

"""
Self-hosted MinerU client for the `/file_parse` endpoint.

This is used by `pdf_to_pptx.py` when MINERU_API_BASE points to an internal
service (e.g. contains "hpecorp"). It downloads the ZIP response and unpacks it
into the same layout that the cloud MinerU flow produces:

  <upload_folder>/mineru_files/<extract_id>/
    *_content_list.json
    layout.json
    full.md
    images/
    ...
"""

from __future__ import annotations

import datetime as _dt
import os
import uuid
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Dict, Optional, Tuple


DEFAULT_SELF_HOSTED_ENDPOINT = "http://ai23.labs.hpecorp.net:8023/file_parse"


def is_self_hosted_mineru(mineru_api_base: str) -> bool:
    return "hpecorp" in (mineru_api_base or "").lower()


def resolve_self_hosted_endpoint(mineru_api_base_or_endpoint: str) -> str:
    """
    Accepts either:
      - base URL: http://host:port
      - full endpoint: http://host:port/file_parse
    and returns the full /file_parse URL.
    """
    url = (mineru_api_base_or_endpoint or "").strip()
    if not url:
        url = DEFAULT_SELF_HOSTED_ENDPOINT

    if url.endswith("/file_parse"):
        return url
    return url.rstrip("/") + "/file_parse"


def _timestamp() -> str:
    return _dt.datetime.now().strftime("%Y%m%d-%H%M%S")


def _safe_extractall(zf: zipfile.ZipFile, out_dir: Path) -> None:
    """
    Prevent Zip Slip by ensuring every extracted path stays within out_dir.
    """
    out_dir_resolved = out_dir.resolve()
    for member in zf.infolist():
        dest = (out_dir_resolved / member.filename).resolve()
        if not str(dest).startswith(str(out_dir_resolved) + os.sep) and dest != out_dir_resolved:
            raise RuntimeError(f"Unsafe path in zip: {member.filename}")
    zf.extractall(out_dir_resolved)


def _flatten_single_top_level_dir(extract_dir: Path) -> None:
    """
    If the zip extracted into a single top-level directory (common pattern),
    move its contents up one level so expected files live directly under extract_dir.
    """
    # Ignore artifacts we may have created (saved zip) and common zip metadata dirs.
    # In our flow, extract_dir usually contains:
    #   - <something>.zip (saved response)
    #   - <top_level_dir>/... (actual extracted content)
    ignored_names = {"__MACOSX", ".DS_Store"}
    candidates = []
    for p in extract_dir.iterdir():
        if p.name in ignored_names:
            continue
        if p.is_file() and p.suffix.lower() == ".zip":
            continue
        candidates.append(p)

    if len(candidates) != 1:
        return

    only = candidates[0]
    if not only.is_dir():
        return

    # Move every entry in <extract_dir>/<only>/ up to <extract_dir>/
    for p in only.iterdir():
        target = extract_dir / p.name
        if target.exists():
            raise RuntimeError(f"Cannot flatten zip output; target already exists: {target}")
        p.rename(target)
    # Remove the now-empty directory (best-effort)
    try:
        only.rmdir()
    except OSError:
        pass


def _ensure_layout_json_from_middle_json(*, extract_dir: Path, pdf_stem: str) -> None:
    """
    Self-hosted MinerU may emit `<name>_middle.json` instead of `layout.json`.
    To keep downstream code unchanged, materialize `layout.json` by copying
    the appropriate `*_middle.json` when `layout.json` is missing.
    """
    layout_path = extract_dir / "layout.json"
    if layout_path.exists():
        return

    preferred = extract_dir / f"{pdf_stem}_middle.json"
    middle_candidates = list(extract_dir.glob("*_middle.json"))

    chosen: Optional[Path] = None
    if preferred.exists():
        chosen = preferred
    elif len(middle_candidates) == 1:
        chosen = middle_candidates[0]

    if chosen is None:
        if not middle_candidates:
            return  # nothing we can do; caller may choose to proceed without layout.json
        candidates = "\n".join(sorted([p.name for p in middle_candidates]))
        raise RuntimeError(
            "Multiple *_middle.json files found but layout.json is missing. "
            f"Cannot choose which to copy in {extract_dir}.\nCandidates:\n{candidates}"
        )

    layout_path.write_bytes(chosen.read_bytes())


def _post_file_parse(
    *,
    endpoint: str,
    pdf_path: Path,
    form_fields: Dict[str, str],
    timeout_s: int,
) -> bytes:
    import requests  # requests is already used elsewhere in this repo

    with pdf_path.open("rb") as f:
        files = {"files": (pdf_path.name, f, "application/pdf")}
        resp = requests.post(endpoint, data=form_fields, files=files, timeout=timeout_s)

    if not resp.ok:
        # Preserve payload in exception for debugging (truncated)
        snippet = resp.text
        if len(snippet) > 10_000:
            snippet = snippet[:10_000] + "\n... (truncated) ..."
        raise RuntimeError(f"Self-hosted MinerU error HTTP {resp.status_code}: {snippet}")

    return resp.content


def parse_pdf_via_self_hosted_mineru(
    pdf_path: Path,
    *,
    mineru_api_base: str,
    upload_folder: Path,
    output_dir_field: str = "./output",
    timeout_s: int = 300,
    form_overrides: Optional[Dict[str, str]] = None,
) -> Tuple[str, Path]:
    """
    Parse a PDF via the self-hosted MinerU `/file_parse` API, save/unzip the ZIP
    response into <upload_folder>/mineru_files/<extract_id>/, and return:
      (extract_id, extract_dir)
    """
    endpoint = resolve_self_hosted_endpoint(mineru_api_base)
    pdf_path = pdf_path.expanduser().resolve()
    upload_folder = upload_folder.expanduser().resolve()

    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    extract_id = uuid.uuid4().hex[:8]
    extract_dir = upload_folder / "mineru_files" / extract_id
    extract_dir.mkdir(parents=True, exist_ok=True)

    # Mirror scratchpad/test_mineru_curl.sh defaults
    fields: Dict[str, str] = {
        "return_middle_json": "true",
        "return_model_output": "true",
        "return_md": "true",
        "return_images": "true",
        "end_page_id": "99999",
        "parse_method": "auto",
        "start_page_id": "0",
        "lang_list": "en",
        "output_dir": output_dir_field,
        "server_url": "string",
        "return_content_list": "true",
        "backend": "hybrid-auto-engine",
        "table_enable": "true",
        "response_format_zip": "true",
        "formula_enable": "true",
    }
    if form_overrides:
        fields.update({k: str(v) for k, v in form_overrides.items()})

    content = _post_file_parse(
        endpoint=endpoint,
        pdf_path=pdf_path,
        form_fields=fields,
        timeout_s=timeout_s,
    )

    zip_path = extract_dir / f"{pdf_path.stem}-mineru-{_timestamp()}.zip"
    zip_path.write_bytes(content)

    if not zipfile.is_zipfile(BytesIO(content)):
        sniff = content[:5000].decode("utf-8", errors="replace")
        raise RuntimeError(
            "Self-hosted MinerU returned non-zip payload. "
            f"Saved to {zip_path}. Snippet:\n{sniff}"
        )

    with zipfile.ZipFile(zip_path, "r") as zf:
        _safe_extractall(zf, extract_dir)

    _flatten_single_top_level_dir(extract_dir)
    _ensure_layout_json_from_middle_json(extract_dir=extract_dir, pdf_stem=pdf_path.stem)

    # Validate expected artifacts exist
    content_list_files = list(extract_dir.glob("*_content_list.json"))
    if not content_list_files:
        # Provide a short directory listing to aid debugging
        entries = sorted([p.relative_to(extract_dir).as_posix() for p in extract_dir.rglob("*")])
        preview = "\n".join(entries[:200])
        if len(entries) > 200:
            preview += "\n... (truncated) ..."
        raise RuntimeError(
            "Self-hosted MinerU ZIP extracted, but no *_content_list.json was found under "
            f"{extract_dir}.\nExtracted entries:\n{preview}"
        )

    return extract_id, extract_dir



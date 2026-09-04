#!/usr/bin/env python3
"""Recursively convert PDFs to Markdown with MinerU.

The standard MinerU API accepts local files through a signed upload URL and
has a 200 MB / 200 page limit.  This tool streams files to the upload URL and
automatically splits larger PDFs into page-based chunks before merging the
returned Markdown and images.
"""

from __future__ import annotations

import argparse
import hashlib
import ipaddress
import json
import logging
import os
import random
import re
import shutil
import socket
import sys
import tempfile
import time
import traceback
import urllib.parse
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Iterator, Mapping

import requests
import urllib3
from pypdf import PdfReader, PdfWriter


LOGGER = logging.getLogger("mineru_pdf_to_md")
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".jp2", ".webp", ".gif", ".bmp", ".svg"}
MARKDOWN_IMAGE_RE = re.compile(r"(!\[[^\]]*\]\()([^\)\r\n]+)(\))")
INVALID_WINDOWS_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED_WINDOWS_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


class MineruError(RuntimeError):
    """Base error for expected conversion failures."""


class MineruApiError(MineruError):
    def __init__(
        self,
        message: str,
        *,
        retryable: bool = False,
        api_code: str | int | None = None,
        http_status: int | None = None,
    ) -> None:
        super().__init__(message)
        self.retryable = retryable
        self.api_code = api_code
        self.http_status = http_status


class PollTimeoutError(MineruApiError):
    def __init__(self, message: str) -> None:
        super().__init__(message, retryable=True)


class LocalPdfError(MineruError):
    pass


@dataclass(frozen=True)
class PdfItem:
    source: Path
    relative_path: Path
    output_dir: Path


def atomic_write_json(path: Path, payload: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        Path(temp_name).replace(path)
    finally:
        temp_path = Path(temp_name)
        if temp_path.exists():
            temp_path.unlink()


def read_json(path: Path) -> dict[str, Any] | None:
    try:
        with path.open("r", encoding="utf-8") as handle:
            value = json.load(handle)
        return value if isinstance(value, dict) else None
    except (OSError, json.JSONDecodeError):
        return None


def safe_name(value: str, fallback: str = "document") -> str:
    value = INVALID_WINDOWS_CHARS_RE.sub("_", value).rstrip(" .")
    if not value:
        value = fallback
    if value.upper() in RESERVED_WINDOWS_NAMES:
        value = f"_{value}"
    return value


def redact(value: Any, secret: str | None = None, limit: int = 1000) -> str:
    text = str(value)
    if secret:
        text = text.replace(secret, "***")
    text = re.sub(r"(Bearer\s+)[^\s,;]+", r"\1***", text, flags=re.IGNORECASE)
    return text[:limit]


def load_simple_dotenv(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.is_file():
        return values
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError:
        return values
    for line in lines:
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
            value = value[1:-1]
        if key:
            values[key] = value
    return values


def application_dir() -> Path:
    """Return the directory containing this tool, independent of cwd."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def find_token(args: argparse.Namespace, input_dir: Path) -> str | None:
    if args.token:
        return args.token.strip()
    if args.token_file:
        try:
            return Path(args.token_file).read_text(encoding="utf-8").strip()
        except OSError as exc:
            raise MineruError(f"无法读取 Token 文件: {args.token_file}: {exc}") from exc

    # File-based configuration is the default-friendly option.  The file must
    # contain only the MinerU token, with optional surrounding whitespace.
    token_file_candidates = [
        Path.cwd() / "mineru_token.txt",
        application_dir() / "mineru_token.txt",
        input_dir / "mineru_token.txt",
    ]
    for token_file in token_file_candidates:
        if token_file.is_file():
            try:
                token = token_file.read_text(encoding="utf-8").strip()
            except OSError as exc:
                raise MineruError(f"无法读取 Token 文件: {token_file}: {exc}") from exc
            if token:
                return token

    for key in ("MINERU_TOKEN", "MINERU_API_TOKEN"):
        if os.environ.get(key, "").strip():
            return os.environ[key].strip()

    dotenv_values: dict[str, str] = {}
    candidates = [Path.cwd() / ".env", application_dir() / ".env", input_dir / ".env"]
    for candidate in candidates:
        dotenv_values.update(load_simple_dotenv(candidate))
    for key in ("MINERU_TOKEN", "MINERU_API_TOKEN"):
        if dotenv_values.get(key, "").strip():
            return dotenv_values[key].strip()
    return None


def is_retryable_api_code(code: Any) -> bool:
    return str(code) in {
        "-10001",  # service exception
        "-60001",  # upload URL generation failure
        "-60007",  # model service temporarily unavailable
        "-60009",  # queue full
        "-60010",  # parse failure that may be transient
        "-60020",  # split failure
        "-60021",  # page count read failure
        "-60022",  # web/network read failure
    }


def is_retryable_http_status(status: int) -> bool:
    return status in {408, 425, 429} or 500 <= status <= 599


def public_ipv4_answers(payload: Mapping[str, Any]) -> list[str]:
    """Extract public IPv4 addresses from a DNS-over-HTTPS JSON response."""
    values: list[str] = []
    answers = payload.get("Answer") or []
    if not isinstance(answers, list):
        return values
    for answer in answers:
        if not isinstance(answer, Mapping) or answer.get("type") != 1:
            continue
        value = str(answer.get("data", "")).strip()
        try:
            address = ipaddress.ip_address(value)
        except ValueError:
            continue
        if address.version == 4 and address.is_global and value not in values:
            values.append(value)
    return values


def is_rfc2544_fake_ip(value: str) -> bool:
    """Return whether an address is in the 198.18.0.0/15 benchmark range."""
    try:
        return ipaddress.ip_address(value) in ipaddress.ip_network("198.18.0.0/15")
    except ValueError:
        return False


class MineruApi:
    def __init__(
        self,
        token: str,
        *,
        base_url: str = "https://mineru.net",
        model_version: str = "vlm",
        language: str = "en",
        enable_table: bool = True,
        enable_formula: bool = True,
        is_ocr: bool = False,
        retry_times: int = 5,
        retry_base_seconds: float = 2.0,
        upload_timeout: float = 3600.0,
        download_timeout: float = 3600.0,
        session: requests.Session | None = None,
    ) -> None:
        self.token = token
        self.base_url = base_url.rstrip("/")
        self.model_version = model_version
        self.language = language
        self.enable_table = enable_table
        self.enable_formula = enable_formula
        self.is_ocr = is_ocr
        self.retry_times = max(1, retry_times)
        self.retry_base_seconds = max(0.0, retry_base_seconds)
        self.upload_timeout = upload_timeout
        self.download_timeout = download_timeout
        self.session = session or requests.Session()

    def _auth_headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self.token}",
            "Accept": "application/json",
        }

    def _backoff(self, attempt: int) -> None:
        if self.retry_base_seconds <= 0:
            return
        delay = min(60.0, self.retry_base_seconds * (2 ** max(0, attempt - 1)))
        delay *= 0.85 + random.random() * 0.3
        LOGGER.warning("API 暂时不可用，将在 %.1f 秒后重试（第 %d/%d 次）", delay, attempt, self.retry_times)
        time.sleep(delay)

    def _request_json(
        self,
        method: str,
        url: str,
        *,
        json_body: Mapping[str, Any] | None = None,
        timeout: tuple[float, float] | float = (30.0, 120.0),
    ) -> dict[str, Any]:
        last_error: Exception | None = None
        for attempt in range(1, self.retry_times + 1):
            try:
                response = self.session.request(
                    method,
                    url,
                    headers=self._auth_headers(),
                    json=json_body,
                    timeout=timeout,
                )
                if response.status_code >= 400:
                    body = redact(response.text, self.token)
                    retryable = is_retryable_http_status(response.status_code)
                    error = MineruApiError(
                        f"HTTP {response.status_code}: {body}",
                        retryable=retryable,
                        http_status=response.status_code,
                    )
                    if retryable and attempt < self.retry_times:
                        self._backoff(attempt)
                        continue
                    raise error

                try:
                    result = response.json()
                except ValueError as exc:
                    last_error = MineruApiError(
                        f"API 返回不是合法 JSON: {redact(response.text, self.token)}",
                        retryable=True,
                        http_status=response.status_code,
                    )
                    if attempt < self.retry_times:
                        self._backoff(attempt)
                        continue
                    raise last_error from exc

                if not isinstance(result, dict):
                    raise MineruApiError("API 返回结构不是 JSON 对象", retryable=False)
                code = result.get("code", 0)
                if str(code) not in {"0", "None"}:
                    message = redact(result.get("msg", "未知 API 错误"), self.token)
                    error = MineruApiError(
                        f"MinerU API 错误 code={code}: {message}",
                        retryable=is_retryable_api_code(code),
                        api_code=code,
                        http_status=response.status_code,
                    )
                    if error.retryable and attempt < self.retry_times:
                        self._backoff(attempt)
                        continue
                    raise error
                return result
            except MineruApiError:
                raise
            except (requests.RequestException, OSError) as exc:
                last_error = exc
                if attempt < self.retry_times:
                    LOGGER.warning("网络请求异常: %s", redact(exc, self.token))
                    self._backoff(attempt)
                    continue
                raise MineruApiError(
                    f"网络请求失败（已重试 {self.retry_times} 次）: {redact(exc, self.token)}",
                    retryable=True,
                ) from exc
        raise MineruApiError(f"API 请求失败: {redact(last_error, self.token)}", retryable=True)

    def create_upload_task(self, file_name: str, data_id: str) -> tuple[str, str]:
        payload = {
            "files": [{"name": file_name, "data_id": data_id, "is_ocr": self.is_ocr}],
            "model_version": self.model_version,
            "enable_formula": self.enable_formula,
            "enable_table": self.enable_table,
            "language": self.language,
        }
        result = self._request_json(
            "POST",
            f"{self.base_url}/api/v4/file-urls/batch",
            json_body=payload,
        )
        data = result.get("data") or {}
        batch_id = data.get("batch_id")
        urls = data.get("file_urls")
        if not batch_id or not isinstance(urls, list) or not urls or not urls[0]:
            raise MineruApiError("申请 MinerU 上传地址成功但响应缺少 batch_id/file_urls", retryable=False)
        return str(batch_id), str(urls[0])

    def upload_file(self, file_path: Path, upload_url: str) -> None:
        last_error: Exception | None = None
        for attempt in range(1, self.retry_times + 1):
            try:
                with file_path.open("rb") as handle:
                    response = self.session.put(
                        upload_url,
                        data=handle,
                        # MinerU explicitly says not to set Content-Type for this PUT.
                        timeout=(30.0, self.upload_timeout),
                    )
                if response.status_code in {200, 201, 204}:
                    return
                # A signed OSS URL can return 401/403 when it expires.  The
                # caller will request a fresh signed URL on the next task
                # retry instead of treating this as a permanent parse error.
                retryable = is_retryable_http_status(response.status_code) or response.status_code in {401, 403}
                error = MineruApiError(
                    f"文件上传失败 HTTP {response.status_code}: {redact(response.text, self.token)}",
                    retryable=retryable,
                    http_status=response.status_code,
                )
                if retryable and attempt < self.retry_times:
                    self._backoff(attempt)
                    continue
                raise error
            except MineruApiError:
                raise
            except (requests.RequestException, OSError) as exc:
                last_error = exc
                if attempt < self.retry_times:
                    LOGGER.warning("文件上传网络异常: %s", redact(exc, self.token))
                    self._backoff(attempt)
                    continue
                raise MineruApiError(
                    f"文件上传失败（已重试 {self.retry_times} 次）: {redact(exc, self.token)}",
                    retryable=True,
                ) from exc
        raise MineruApiError(f"文件上传失败: {redact(last_error, self.token)}", retryable=True)

    def get_batch_result(self, batch_id: str) -> dict[str, Any]:
        result = self._request_json(
            "GET",
            f"{self.base_url}/api/v4/extract-results/batch/{urllib.parse.quote(batch_id, safe="")}",
        )
        data = result.get("data") or {}
        entries = data.get("extract_result") or []
        if isinstance(entries, dict):
            entries = [entries]
        if not isinstance(entries, list) or not entries:
            raise MineruApiError("任务查询响应缺少 extract_result", retryable=True)
        entry = entries[0]
        if not isinstance(entry, dict):
            raise MineruApiError("任务查询响应中的 extract_result 格式错误", retryable=False)
        return entry

    def wait_for_batch(
        self,
        batch_id: str,
        *,
        file_name: str,
        data_id: str,
        poll_interval: float,
        poll_timeout: float,
    ) -> dict[str, Any]:
        started = time.monotonic()
        last_state: str | None = None
        while time.monotonic() - started <= poll_timeout:
            entry = self.get_batch_result(batch_id)
            state = str(entry.get("state", "unknown"))
            progress = entry.get("extract_progress") or {}
            progress_text = ""
            if isinstance(progress, dict) and progress.get("total_pages"):
                progress_text = f" {progress.get('extracted_pages', 0)}/{progress.get('total_pages')} 页"
            if state != last_state or progress_text:
                LOGGER.info("%s: %s%s", file_name, state, progress_text)
                last_state = state
            if state == "done":
                url = entry.get("full_zip_url")
                if not url:
                    raise MineruApiError("任务完成但响应缺少 full_zip_url", retryable=True)
                return entry
            if state == "failed":
                error_code = entry.get("err_code")
                message = entry.get("err_msg") or "未知解析失败"
                raise MineruApiError(
                    f"文件解析失败 code={error_code}: {redact(message, self.token)}",
                    retryable=is_retryable_api_code(error_code),
                    api_code=error_code,
                )
            time.sleep(max(0.1, poll_interval))
        raise PollTimeoutError(f"轮询任务超时（{poll_timeout:.0f} 秒）: {batch_id} ({file_name})")

    def download_zip(self, url: str, destination: Path) -> None:
        destination.parent.mkdir(parents=True, exist_ok=True)
        last_error: Exception | None = None
        parsed = urllib.parse.urlsplit(url)
        hostname = parsed.hostname or ""
        if hostname == "cdn-mineru.openxlab.org.cn":
            try:
                local_addresses = {
                    item[4][0]
                    for item in socket.getaddrinfo(hostname, 443, type=socket.SOCK_STREAM)
                }
            except OSError:
                local_addresses = set()
            if any(is_rfc2544_fake_ip(value) for value in local_addresses):
                LOGGER.warning("检测到 MinerU CDN 被解析为 198.18.x.x fake-IP，改用保留 TLS 校验的直连下载")
                try:
                    self._download_zip_via_public_ip(url, destination)
                    return
                except Exception as exc:
                    last_error = exc
                    LOGGER.warning("CDN fake-IP 直连回退失败，将尝试系统网络: %s", redact(exc, self.token))
        for attempt in range(1, self.retry_times + 1):
            try:
                # The URL is already signed; do not send the API token to the CDN.
                with self.session.get(
                    url,
                    headers={"Accept": "application/zip, application/octet-stream, */*"},
                    stream=True,
                    timeout=(30.0, self.download_timeout),
                ) as response:
                    if response.status_code >= 400:
                        retryable = is_retryable_http_status(response.status_code)
                        error = MineruApiError(
                            f"结果 ZIP 下载失败 HTTP {response.status_code}: {redact(response.text, self.token)}",
                            retryable=retryable,
                            http_status=response.status_code,
                        )
                        if retryable and attempt < self.retry_times:
                            self._backoff(attempt)
                            continue
                        raise error
                    with destination.open("wb") as handle:
                        for block in response.iter_content(chunk_size=1024 * 1024):
                            if block:
                                handle.write(block)
                if destination.stat().st_size == 0:
                    raise MineruApiError("结果 ZIP 为空", retryable=True)
                return
            except MineruApiError:
                raise
            except (requests.RequestException, OSError) as exc:
                last_error = exc
                if attempt < self.retry_times:
                    LOGGER.warning("结果 ZIP 下载网络异常: %s", redact(exc, self.token))
                    self._backoff(attempt)
                    continue
                break

        # Clash and similar TUN proxies may return an RFC 2544 fake IP
        # (198.18.0.0/15) for the MinerU CDN.  In that failure mode the API and
        # upload succeed but the CDN TLS handshake ends with UNEXPECTED_EOF.
        # Resolve the current public address through DoH, connect to that IP,
        # and still use the original hostname for SNI and certificate checks.
        try:
            self._download_zip_via_public_ip(url, destination)
            return
        except Exception as direct_error:
            raise MineruApiError(
                "结果 ZIP 下载失败；标准连接错误："
                f"{redact(last_error, self.token)}；保留 TLS 校验的直连回退也失败："
                f"{redact(direct_error, self.token)}",
                retryable=True,
            ) from direct_error

    def _download_zip_via_public_ip(self, url: str, destination: Path) -> None:
        parsed = urllib.parse.urlsplit(url)
        hostname = parsed.hostname or ""
        if parsed.scheme != "https" or hostname != "cdn-mineru.openxlab.org.cn":
            raise MineruApiError(f"不允许对非 MinerU CDN 地址执行直连回退: {hostname}")

        addresses: list[str] = []
        override = os.environ.get("MINERU_CDN_IP", "").strip()
        if override:
            try:
                address = ipaddress.ip_address(override)
            except ValueError as exc:
                raise MineruApiError("MINERU_CDN_IP 不是有效的 IP 地址") from exc
            if address.version != 4 or not address.is_global:
                raise MineruApiError("MINERU_CDN_IP 必须是公开 IPv4 地址")
            addresses.append(override)
        else:
            response = self.session.get(
                "https://dns.alidns.com/resolve",
                params={"name": hostname, "type": "A"},
                headers={"Accept": "application/dns-json"},
                timeout=(20.0, 60.0),
            )
            response.raise_for_status()
            addresses = public_ipv4_answers(response.json())
        if not addresses:
            raise MineruApiError("公共 DNS 没有返回 MinerU CDN 的公开 IPv4 地址")

        request_target = urllib.parse.urlunsplit(("", "", parsed.path or "/", parsed.query, ""))
        errors: list[str] = []
        partial = destination.with_suffix(destination.suffix + ".part")
        partial.unlink(missing_ok=True)
        for address in addresses:
            pool = urllib3.HTTPSConnectionPool(
                address,
                port=parsed.port or 443,
                assert_hostname=hostname,
                server_hostname=hostname,
                timeout=urllib3.Timeout(connect=30.0, read=self.download_timeout),
                retries=False,
            )
            response = None
            try:
                response = pool.request(
                    "GET",
                    request_target,
                    headers={"Host": hostname, "Accept": "application/zip, application/octet-stream, */*"},
                    preload_content=False,
                    redirect=False,
                )
                if response.status < 200 or response.status >= 300:
                    raise MineruApiError(f"CDN 直连 HTTP {response.status}")
                with partial.open("wb") as handle:
                    while True:
                        block = response.read(1024 * 1024)
                        if not block:
                            break
                        handle.write(block)
                if partial.stat().st_size == 0:
                    raise MineruApiError("CDN 直连结果 ZIP 为空")
                partial.replace(destination)
                LOGGER.info("MinerU CDN fake-IP 回退成功（TLS 域名校验保持开启）")
                return
            except Exception as exc:
                partial.unlink(missing_ok=True)
                errors.append(f"{address}: {redact(exc, self.token, limit=300)}")
            finally:
                if response is not None:
                    response.release_conn()
                pool.close()
        raise MineruApiError("；".join(errors), retryable=True)


def iter_pdfs(input_dir: Path, output_root: Path) -> Iterator[Path]:
    input_dir = input_dir.resolve()
    output_root = output_root.resolve()
    for path in sorted(input_dir.rglob("*"), key=lambda item: str(item).lower()):
        if not path.is_file() or path.suffix.lower() != ".pdf":
            continue
        try:
            if path.resolve().is_relative_to(output_root):
                continue
        except AttributeError:  # pragma: no cover - Python 3.8 compatibility
            if str(path.resolve()).startswith(str(output_root)):
                continue
        yield path


def make_pdf_items(input_dir: Path, output_root: Path) -> list[PdfItem]:
    items: list[PdfItem] = []
    for source in iter_pdfs(input_dir, output_root):
        relative = source.relative_to(input_dir)
        output_dir = output_root / relative.parent / safe_name(source.stem)
        items.append(PdfItem(source=source, relative_path=relative, output_dir=output_dir))
    return items


def default_output_dir(input_dir: Path) -> Path:
    """Return an output directory beside the source directory."""
    input_dir = input_dir.resolve()
    return input_dir.parent / f"{input_dir.name}_mineru_md"


def pdf_signature(source: Path) -> dict[str, Any]:
    stat = source.stat()
    return {
        "path": str(source.resolve()),
        "size": stat.st_size,
        "mtime_ns": stat.st_mtime_ns,
    }


def make_data_id(source: Path, chunk_index: int) -> str:
    raw = f"{source.resolve()}::{chunk_index}".encode("utf-8", errors="replace")
    return f"mineru_{hashlib.sha256(raw).hexdigest()[:32]}"


def write_pages(pages: list[Any], destination: Path) -> None:
    writer = PdfWriter()
    for page in pages:
        writer.add_page(page)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("wb") as handle:
        writer.write(handle)


def split_pdf(
    source: Path,
    work_dir: Path,
    *,
    max_bytes: int,
    max_pages: int,
) -> list[Path]:
    try:
        source_size = source.stat().st_size
        reader = PdfReader(str(source), strict=False)
        page_count = len(reader.pages)
    except Exception as exc:  # pypdf has several exception types across versions
        raise LocalPdfError(f"无法读取 PDF 页数: {source}: {exc}") from exc

    if page_count <= 0:
        raise LocalPdfError(f"PDF 没有页面: {source}")
    if source_size <= max_bytes and page_count <= max_pages:
        return [source]

    work_dir.mkdir(parents=True, exist_ok=True)
    pages = list(reader.pages)
    chunks: list[Path] = []
    offset = 0
    part_number = 1
    while offset < len(pages):
        upper = min(max_pages, len(pages) - offset)
        trial = work_dir / f".part_{part_number:04d}.trial.pdf"
        candidate_len = upper
        write_pages(pages[offset : offset + candidate_len], trial)
        if trial.stat().st_size > max_bytes:
            low, high = 1, candidate_len - 1
            best: int | None = None
            while low <= high:
                middle = (low + high) // 2
                write_pages(pages[offset : offset + middle], trial)
                if trial.stat().st_size <= max_bytes:
                    best = middle
                    low = middle + 1
                else:
                    high = middle - 1
            if best is None:
                trial.unlink(missing_ok=True)
                raise LocalPdfError(
                    f"PDF 的单页超过 MinerU 安全上传上限 {max_bytes / (1024 ** 2):.1f} MiB，无法自动拆分: {source}"
                )
            candidate_len = best
            write_pages(pages[offset : offset + candidate_len], trial)

        final_path = work_dir / f"part_{part_number:04d}.pdf"
        trial.replace(final_path)
        chunks.append(final_path)
        offset += candidate_len
        part_number += 1
    return chunks


def safe_extract_zip(zip_path: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    root = destination.resolve()
    try:
        with zipfile.ZipFile(zip_path) as archive:
            for member in archive.infolist():
                member_path = (destination / member.filename).resolve()
                if not member_path.is_relative_to(root):
                    raise MineruError(f"结果 ZIP 含有越界路径，已拒绝解压: {member.filename}")
                if member.is_dir():
                    member_path.mkdir(parents=True, exist_ok=True)
                    continue
                member_path.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(member) as source, member_path.open("wb") as target:
                    shutil.copyfileobj(source, target, length=1024 * 1024)
    except zipfile.BadZipFile as exc:
        raise MineruError(f"MinerU 返回的结果不是有效 ZIP: {zip_path}") from exc


def find_markdown(extracted_root: Path) -> Path:
    candidates = [
        path for path in extracted_root.rglob("*") if path.is_file() and path.name.lower() == "full.md"
    ]
    if not candidates:
        candidates = [path for path in extracted_root.rglob("*") if path.is_file() and path.suffix.lower() == ".md"]
    if not candidates:
        raise MineruError(f"结果 ZIP 中没有找到 Markdown 文件: {extracted_root}")
    return sorted(candidates, key=lambda item: len(item.parts))[0]


def unique_image_destination(image_root: Path, prefix: str, original_name: str, used: set[str]) -> Path:
    stem = safe_name(Path(original_name).stem, fallback="image")
    suffix = Path(original_name).suffix.lower()
    base = f"{prefix}{stem}{suffix}"
    candidate = base
    counter = 2
    while candidate.lower() in used or (image_root / candidate).exists():
        candidate = f"{prefix}{stem}_{counter}{suffix}"
        counter += 1
    used.add(candidate.lower())
    return image_root / candidate


def materialize_markdown(
    markdown_path: Path,
    extracted_root: Path,
    output_dir: Path,
    *,
    image_prefix: str,
) -> str:
    image_root = output_dir / "images"
    image_root.mkdir(parents=True, exist_ok=True)
    used_names: set[str] = set()
    resolved_map: dict[Path, Path] = {}
    image_files = sorted(
        (path for path in extracted_root.rglob("*") if path.is_file() and path.suffix.lower() in IMAGE_SUFFIXES),
        key=lambda item: str(item).lower(),
    )
    for image_file in image_files:
        destination = unique_image_destination(image_root, image_prefix, image_file.name, used_names)
        shutil.copyfile(image_file, destination)
        resolved_map[image_file.resolve()] = destination

    try:
        text = markdown_path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        text = markdown_path.read_text(encoding="utf-8", errors="replace")

    by_basename: dict[str, list[Path]] = {}
    for original, destination in resolved_map.items():
        by_basename.setdefault(original.name.lower(), []).append(destination)

    def resolve_image(raw_target: str) -> Path | None:
        target = raw_target.strip()
        if target.startswith("<") and ">" in target:
            target = target[1 : target.find(">")]
        else:
            target = target.split()[0] if target.split() else ""
        parsed = urllib.parse.urlparse(target)
        if parsed.scheme or parsed.netloc or target.startswith("data:"):
            return None
        decoded = urllib.parse.unquote(parsed.path or target).replace("\\", "/")
        candidates = [
            (markdown_path.parent / decoded).resolve(),
            (extracted_root / decoded.lstrip("/\\")).resolve(),
        ]
        for candidate in candidates:
            if candidate in resolved_map:
                return resolved_map[candidate]
        matches = by_basename.get(Path(decoded).name.lower(), [])
        return matches[0] if matches else None

    def replace_image(match: re.Match[str]) -> str:
        destination = resolve_image(match.group(2))
        if destination is None:
            return match.group(0)
        relative_target = Path("images") / destination.name
        return f"{match.group(1)}{relative_target.as_posix()}{match.group(3)}"

    return MARKDOWN_IMAGE_RE.sub(replace_image, text).rstrip() + "\n"


def merge_result_archives(
    archives: list[Path],
    output_dir: Path,
    *,
    output_markdown_name: str,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    image_root = output_dir / "images"
    if image_root.exists():
        shutil.rmtree(image_root)
    image_root.mkdir(parents=True, exist_ok=True)
    parts: list[str] = []
    # Keep extraction in a deterministic, tool-owned directory.  This also
    # makes interrupted runs easy to clean and avoids leaving random temp
    # folders beside the user's Markdown output.
    temp_root = output_dir / ".mineru_extract"
    if temp_root.exists():
        shutil.rmtree(temp_root)
    temp_root.mkdir(parents=True, exist_ok=True)
    try:
        for index, archive in enumerate(archives, start=1):
            extracted = temp_root / f"part_{index:04d}"
            safe_extract_zip(archive, extracted)
            markdown_path = find_markdown(extracted)
            prefix = "" if len(archives) == 1 else f"part_{index:04d}_"
            parts.append(
                materialize_markdown(
                    markdown_path,
                    extracted,
                    output_dir,
                    image_prefix=prefix,
                )
            )
    finally:
        if temp_root.exists():
            shutil.rmtree(temp_root)
    # API chunking is an implementation detail.  Do not insert a Markdown
    # horizontal rule that did not exist in the source document.
    merged = "\n\n".join(parts).rstrip() + "\n"
    markdown_output = output_dir / output_markdown_name
    markdown_output.write_text(merged, encoding="utf-8", newline="\n")
    return markdown_output


def initialize_state(source: Path, chunks: list[Path]) -> dict[str, Any]:
    return {
        "version": 1,
        "source": pdf_signature(source),
        "status": "processing",
        "chunks": [
            {
                "index": index,
                "path": str(chunk.resolve()),
                "file_name": f"{safe_name(source.stem)}_part_{index:04d}.pdf",
                "data_id": make_data_id(source, index),
                "status": "pending",
            }
            for index, chunk in enumerate(chunks, start=1)
        ],
    }


def state_matches_source(state: Mapping[str, Any], source: Path) -> bool:
    saved = state.get("source")
    return isinstance(saved, dict) and saved == pdf_signature(source)


class PdfProcessor:
    def __init__(
        self,
        api: MineruApi,
        *,
        max_bytes: int,
        max_pages: int,
        poll_interval: float,
        poll_timeout: float,
        task_retries: int,
        retry_base_seconds: float,
    ) -> None:
        self.api = api
        self.max_bytes = max_bytes
        self.max_pages = max_pages
        self.poll_interval = poll_interval
        self.poll_timeout = poll_timeout
        self.task_retries = max(1, task_retries)
        self.retry_base_seconds = max(0.0, retry_base_seconds)

    def _retry_delay(self, attempt: int) -> None:
        if self.retry_base_seconds <= 0:
            return
        delay = min(60.0, self.retry_base_seconds * (2 ** max(0, attempt - 1)))
        LOGGER.warning("分片任务将重试，等待 %.1f 秒（第 %d/%d 次）", delay, attempt, self.task_retries)
        time.sleep(delay)

    def _process_chunk(
        self,
        chunk: dict[str, Any],
        state: dict[str, Any],
        state_path: Path,
    ) -> Path:
        chunk_path = Path(str(chunk["path"]))
        if not chunk_path.is_file():
            raise LocalPdfError(f"找不到待上传的分片: {chunk_path}")
        work_dir = state_path.parent / ".mineru_work"
        zip_path = work_dir / f"result_{int(chunk['index']):04d}.zip"
        if chunk.get("status") == "done" and zip_path.is_file() and zip_path.stat().st_size > 0:
            return zip_path

        for attempt in range(1, self.task_retries + 1):
            try:
                batch_id = chunk.get("batch_id")
                upload_url = chunk.get("upload_url")
                if chunk.get("status") in {None, "pending", "retry"} or not batch_id:
                    batch_id, upload_url = self.api.create_upload_task(
                        str(chunk["file_name"]), str(chunk["data_id"])
                    )
                    chunk.update({"batch_id": batch_id, "upload_url": upload_url, "status": "uploading"})
                    atomic_write_json(state_path, state)

                if chunk.get("status") == "uploading":
                    if not upload_url:
                        raise MineruApiError("任务处于 uploading 状态但没有上传 URL", retryable=False)
                    self.api.upload_file(chunk_path, str(upload_url))
                    chunk["status"] = "uploaded"
                    chunk.pop("upload_url", None)
                    atomic_write_json(state_path, state)

                entry = self.api.wait_for_batch(
                    str(batch_id),
                    file_name=str(chunk["file_name"]),
                    data_id=str(chunk["data_id"]),
                    poll_interval=self.poll_interval,
                    poll_timeout=self.poll_timeout,
                )
                chunk["status"] = "downloading"
                chunk["full_zip_url"] = str(entry["full_zip_url"])
                atomic_write_json(state_path, state)
                self.api.download_zip(str(entry["full_zip_url"]), zip_path)
                chunk["status"] = "done"
                chunk.pop("upload_url", None)
                atomic_write_json(state_path, state)
                return zip_path
            except MineruApiError as exc:
                if not exc.retryable or attempt >= self.task_retries:
                    raise
                # A failed parse task cannot be repaired in-place.  Start a new
                # task on retry; an uploaded waiting-file task may be abandoned.
                chunk.update({"status": "retry", "batch_id": None, "upload_url": None})
                atomic_write_json(state_path, state)
                self._retry_delay(attempt)
        raise MineruApiError(f"分片处理失败: {chunk.get('file_name')}", retryable=True)

    def process(self, item: PdfItem, *, overwrite: bool = False) -> str:
        output_dir = item.output_dir
        output_dir.mkdir(parents=True, exist_ok=True)
        markdown_name = f"{safe_name(item.source.stem)}.md"
        markdown_path = output_dir / markdown_name
        state_path = output_dir / ".mineru_state.json"
        if markdown_path.is_file() and not overwrite:
            return "skipped"

        work_dir = output_dir / ".mineru_work"
        old_state = read_json(state_path) if not overwrite else None
        state = old_state if old_state and state_matches_source(old_state, item.source) else None
        if not state:
            if work_dir.exists():
                shutil.rmtree(work_dir)
            chunks = split_pdf(
                item.source,
                work_dir,
                max_bytes=self.max_bytes,
                max_pages=self.max_pages,
            )
            state = initialize_state(item.source, chunks)
            atomic_write_json(state_path, state)
        else:
            LOGGER.info("发现未完成状态，继续处理: %s", item.source)

        archives: list[Path] = []
        try:
            for chunk in state.get("chunks", []):
                if not isinstance(chunk, dict):
                    raise MineruError("状态文件中的分片记录格式错误")
                archives.append(self._process_chunk(chunk, state, state_path))
            merge_result_archives(archives, output_dir, output_markdown_name=markdown_name)
            state["status"] = "done"
            state["completed_at"] = time.strftime("%Y-%m-%dT%H:%M:%S%z")
            atomic_write_json(
                output_dir / ".mineru_meta.json",
                {
                    "source": state["source"],
                    "status": "done",
                    "chunks": len(archives),
                    "markdown": markdown_name,
                    "images_dir": "images",
                    "model_version": self.api.model_version,
                    "language": self.api.language,
                    "completed_at": state["completed_at"],
                },
            )
            state_path.unlink(missing_ok=True)
            if work_dir.exists():
                shutil.rmtree(work_dir)
            return "converted"
        except Exception as exc:
            state["status"] = "failed"
            state["error"] = redact(exc, self.api.token)
            state["traceback"] = traceback.format_exc(limit=8)
            atomic_write_json(state_path, state)
            raise


def format_bytes(value: int) -> str:
    if value >= 1024**3:
        return f"{value / 1024**3:.2f} GiB"
    return f"{value / 1024**2:.1f} MiB"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="递归调用 MinerU 将 PDF 转为 Markdown，并保留图片集。",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("input_dir", type=Path, help="输入 PDF 文件夹，例如 E:\\Tessent\\9.Gemini\\paper")
    parser.add_argument("--output-dir", type=Path, help="输出根目录；不指定时使用 <输入目录>_mineru_md")
    parser.add_argument("--token", help="MinerU Token；更推荐使用 MINERU_TOKEN 环境变量")
    parser.add_argument(
        "--token-file",
        type=Path,
        help="从文本文件读取 MinerU Token；不指定时自动查找 mineru_token.txt",
    )
    parser.add_argument("--api-base-url", default="https://mineru.net")
    parser.add_argument("--model-version", choices=("pipeline", "vlm"), default="vlm")
    parser.add_argument("--language", default="en", help="文档语言；中文 PDF 可传 ch")
    parser.add_argument("--ocr", action="store_true", help="开启 OCR")
    parser.add_argument("--no-table", action="store_true", help="关闭表格识别")
    parser.add_argument("--no-formula", action="store_true", help="关闭公式识别")
    parser.add_argument(
        "--max-api-file-mb",
        type=float,
        default=190.0,
        help="单个 API 分片的安全大小上限；MinerU 文档上限为 200 MB",
    )
    parser.add_argument(
        "--max-pages-per-chunk",
        type=int,
        default=180,
        help="单个 API 分片的页数上限；MinerU 文档上限为 200 页",
    )
    parser.add_argument("--poll-interval", type=float, default=5.0, help="任务轮询间隔（秒）")
    parser.add_argument("--poll-timeout", type=float, default=21600.0, help="单个分片最大轮询时间（秒）")
    parser.add_argument("--retry-times", type=int, default=5, help="单次 HTTP 操作最大尝试次数")
    parser.add_argument("--task-retries", type=int, default=3, help="解析分片失败时重新提交的次数")
    parser.add_argument("--retry-base-seconds", type=float, default=2.0, help="指数退避初始秒数")
    parser.add_argument("--upload-timeout", type=float, default=3600.0, help="单次文件上传读超时（秒）")
    parser.add_argument("--download-timeout", type=float, default=3600.0, help="结果下载读超时（秒）")
    parser.add_argument("--limit", type=int, help="只处理前 N 个 PDF，适合先试跑")
    parser.add_argument("--overwrite", action="store_true", help="重新转换已有输出")
    parser.add_argument("--dry-run", action="store_true", help="只扫描并显示映射，不调用 API")
    parser.add_argument("--verbose", action="store_true", help="输出更详细日志")
    return parser


def run(args: argparse.Namespace) -> int:
    input_dir = args.input_dir.expanduser().resolve()
    if not input_dir.is_dir():
        raise MineruError(f"输入目录不存在或不是目录: {input_dir}")
    output_root = args.output_dir.expanduser().resolve() if args.output_dir else default_output_dir(input_dir)
    output_root.mkdir(parents=True, exist_ok=True)
    items = make_pdf_items(input_dir, output_root)
    if args.limit is not None:
        if args.limit < 1:
            raise MineruError("--limit 必须大于 0")
        items = items[: args.limit]
    if not items:
        LOGGER.warning("没有找到 PDF: %s", input_dir)
        return 0

    LOGGER.info("输入目录: %s", input_dir)
    LOGGER.info("输出目录: %s", output_root)
    LOGGER.info("发现 %d 个 PDF", len(items))
    if args.dry_run:
        for index, item in enumerate(items, start=1):
            size = format_bytes(item.source.stat().st_size)
            print(f"[{index}/{len(items)}] {item.relative_path} ({size}) -> {item.output_dir}")
        return 0

    token = find_token(args, input_dir)
    if not token:
        raise MineruError(
            "未找到 MinerU Token。请在 mineru_token.txt 中填写 Token，"
            "或使用 --token-file；不建议把 Token 直接写入命令历史。"
        )
    if args.max_api_file_mb <= 0 or args.max_api_file_mb >= 200:
        raise MineruError("--max-api-file-mb 应在 0 和 200 之间，建议使用默认的 190")
    if args.max_pages_per_chunk <= 0 or args.max_pages_per_chunk > 200:
        raise MineruError("--max-pages-per-chunk 应在 1 到 200 之间")

    api = MineruApi(
        token,
        base_url=args.api_base_url,
        model_version=args.model_version,
        language=args.language,
        enable_table=not args.no_table,
        enable_formula=not args.no_formula,
        is_ocr=args.ocr,
        retry_times=args.retry_times,
        retry_base_seconds=args.retry_base_seconds,
        upload_timeout=args.upload_timeout,
        download_timeout=args.download_timeout,
    )
    processor = PdfProcessor(
        api,
        max_bytes=int(args.max_api_file_mb * 1024 * 1024),
        max_pages=args.max_pages_per_chunk,
        poll_interval=args.poll_interval,
        poll_timeout=args.poll_timeout,
        task_retries=args.task_retries,
        retry_base_seconds=args.retry_base_seconds,
    )

    summary: dict[str, Any] = {
        "input_dir": str(input_dir),
        "output_dir": str(output_root),
        "started_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "total": len(items),
        "converted": 0,
        "skipped": 0,
        "failed": 0,
        "files": [],
    }
    for index, item in enumerate(items, start=1):
        LOGGER.info("[%d/%d] 开始: %s", index, len(items), item.relative_path)
        file_record: dict[str, Any] = {
            "source": str(item.source),
            "relative_path": str(item.relative_path),
            "output_dir": str(item.output_dir),
        }
        try:
            status = processor.process(item, overwrite=args.overwrite)
            summary[status] = int(summary.get(status, 0)) + 1
            file_record["status"] = status
            LOGGER.info("[%d/%d] 完成: %s", index, len(items), status)
        except Exception as exc:
            summary["failed"] += 1
            file_record.update({"status": "failed", "error": redact(exc, api.token)})
            LOGGER.error("[%d/%d] 失败: %s: %s", index, len(items), item.relative_path, redact(exc, api.token))
        summary["files"].append(file_record)
        atomic_write_json(output_root / "run_summary.json", summary)

    summary["finished_at"] = time.strftime("%Y-%m-%dT%H:%M:%S%z")
    atomic_write_json(output_root / "run_summary.json", summary)
    LOGGER.info(
        "处理结束：converted=%d skipped=%d failed=%d，详情见 %s",
        summary["converted"],
        summary["skipped"],
        summary["failed"],
        output_root / "run_summary.json",
    )
    return 1 if summary["failed"] else 0


def main(argv: Iterable[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    try:
        return run(args)
    except MineruError as exc:
        LOGGER.error("%s", exc)
        return 2
    except KeyboardInterrupt:
        LOGGER.error("用户中断；未完成任务的 .mineru_state.json 会保留以便下次续跑。")
        return 130


if __name__ == "__main__":
    sys.exit(main())

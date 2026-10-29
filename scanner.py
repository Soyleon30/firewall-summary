from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Optional
import os
import re

try:
    from openpyxl import load_workbook  # type: ignore
except Exception:  # pragma: no cover - optional dependency handled at runtime
    load_workbook = None  # type: ignore

try:
    import xlrd  # type: ignore
except Exception:  # pragma: no cover - optional dependency handled at runtime
    xlrd = None  # type: ignore


SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".xls"}
HEADER_TOKENS = {
    "原地址",
    "源地址",
    "目的地址",
    "目标地址",
    "目的端口",
    "目标端口",
    "端口",
}


ProgressCallback = Callable[[int, int, str], None]
LogCallback = Callable[[str], None]


@dataclass(slots=True)
class RequestRow:
    sheet_name: str
    row_number: int
    source_address: str
    destination_address: str
    destination_port: str


@dataclass(slots=True)
class FirewallRecord:
    date_text: str
    date_sort: Optional[tuple[int, int, int]]
    request_dir: str
    request_dir_rel: str
    file_name: str
    file_path: str
    sheet_name: str
    row_number: int
    source_address: str
    destination_address: str
    destination_port: str
    normalized_key: tuple[str, str, str]
    group_id: int = 0
    group_size: int = 1
    cross_directory_duplicate: bool = False

    def sort_key(self) -> tuple:
        date_key = self.date_sort if self.date_sort else (9999, 99, 99)
        return (
            0 if self.date_sort else 1,
            date_key,
            self.request_dir_rel.casefold(),
            self.file_name.casefold(),
            self.sheet_name.casefold(),
            self.row_number,
            self.source_address.casefold(),
            self.destination_address.casefold(),
            self.destination_port.casefold(),
        )


@dataclass(slots=True)
class ScanSummary:
    root_path: str = ""
    discovered_excel_files: int = 0
    processed_files: int = 0
    matched_directories: int = 0
    parsed_sheets: int = 0
    valid_rows: int = 0
    kept_records: int = 0
    skipped_same_dir_duplicates: int = 0
    skipped_incomplete_rows: int = 0
    skipped_outside_firewall_dirs: int = 0
    skipped_unreadable_files: int = 0
    cross_directory_duplicate_groups: int = 0
    cross_directory_duplicate_records: int = 0
    cancelled: bool = False
    errors: list[str] = field(default_factory=list)


@dataclass(slots=True)
class ScanResult:
    summary: ScanSummary
    records: list[FirewallRecord] = field(default_factory=list)
    groups: list[list[FirewallRecord]] = field(default_factory=list)


def _clean_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        text = "TRUE" if value else "FALSE"
    elif isinstance(value, float):
        if value.is_integer():
            text = str(int(value))
        else:
            text = ("%f" % value).rstrip("0").rstrip(".")
    else:
        text = str(value)
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _parse_date_segment(name: str) -> Optional[tuple[str, tuple[int, int, int]]]:
    patterns = (
        re.compile(r"^(?P<y>20\d{2})[.\-_/年](?P<m>\d{1,2})[.\-_/月](?P<d>\d{1,2})日?"),
        re.compile(r"^(?P<y>20\d{2})(?P<m>\d{2})(?P<d>\d{2})"),
    )
    head = name.strip()
    for pattern in patterns:
        match = pattern.match(head)
        if not match:
            continue
        year = int(match.group("y"))
        month = int(match.group("m"))
        day = int(match.group("d"))
        if 1 <= month <= 12 and 1 <= day <= 31:
            return f"{year:04d}-{month:02d}-{day:02d}", (year, month, day)
    return None


def _extract_date_from_request_dir(request_dir: Path, root_path: Path) -> tuple[str, Optional[tuple[int, int, int]]]:
    current = request_dir
    while True:
        parsed = _parse_date_segment(current.name)
        if parsed:
            return parsed
        if current == root_path or current.parent == current:
            break
        current = current.parent
    return "未识别", None


def _is_header_row(source: str, destination: str, port: str) -> bool:
    values = [source, destination, port]
    filled = [item for item in values if item]
    if not filled:
        return False
    return all(item in HEADER_TOKENS for item in filled)


def _iter_xlsx_rows(file_path: Path) -> tuple[list[RequestRow], int]:
    if load_workbook is None:
        raise RuntimeError("缺少 openpyxl 依赖，无法读取 .xlsx/.xlsm 文件。")

    workbook = load_workbook(file_path, read_only=True, data_only=True)
    rows: list[RequestRow] = []
    sheet_count = 0
    try:
        for worksheet in workbook.worksheets:
            sheet_count += 1
            for row_index, row in enumerate(
                worksheet.iter_rows(min_row=1, max_col=8, values_only=True),
                start=1,
            ):
                source = _clean_text(row[2] if len(row) > 2 else None)
                destination = _clean_text(row[5] if len(row) > 5 else None)
                port = _clean_text(row[7] if len(row) > 7 else None)
                if not (source or destination or port):
                    continue
                if _is_header_row(source, destination, port):
                    continue
                rows.append(
                    RequestRow(
                        sheet_name=worksheet.title,
                        row_number=row_index,
                        source_address=source,
                        destination_address=destination,
                        destination_port=port,
                    )
                )
    finally:
        workbook.close()
    return rows, sheet_count


def _iter_xls_rows(file_path: Path) -> tuple[list[RequestRow], int]:
    if xlrd is None:
        raise RuntimeError("缺少 xlrd 依赖，无法读取 .xls 文件。")

    workbook = xlrd.open_workbook(file_path.as_posix(), on_demand=True)
    rows: list[RequestRow] = []
    sheet_count = 0
    try:
        for worksheet in workbook.sheets():
            sheet_count += 1
            for row_index in range(worksheet.nrows):
                row = worksheet.row_values(row_index)
                source = _clean_text(row[2] if len(row) > 2 else None)
                destination = _clean_text(row[5] if len(row) > 5 else None)
                port = _clean_text(row[7] if len(row) > 7 else None)
                if not (source or destination or port):
                    continue
                if _is_header_row(source, destination, port):
                    continue
                rows.append(
                    RequestRow(
                        sheet_name=worksheet.name,
                        row_number=row_index + 1,
                        source_address=source,
                        destination_address=destination,
                        destination_port=port,
                    )
                )
    finally:
        workbook.release_resources()
    return rows, sheet_count


def _read_workbook_rows(file_path: Path) -> tuple[list[RequestRow], int]:
    suffix = file_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return _iter_xlsx_rows(file_path)
    if suffix == ".xls":
        return _iter_xls_rows(file_path)
    return [], 0


class FirewallScanner:
    def __init__(self, root_path: str | Path) -> None:
        self.root_path = Path(root_path).expanduser()

    def collect_excel_files(self) -> list[Path]:
        files: list[Path] = []
        for dirpath, dirnames, filenames in os.walk(self.root_path):
            dirnames[:] = sorted(name for name in dirnames if not name.startswith("."))
            for filename in sorted(filenames):
                if filename.startswith("~$") or filename.startswith("."):
                    continue
                if Path(filename).suffix.lower() in SUPPORTED_SUFFIXES:
                    files.append(Path(dirpath) / filename)
        return files

    def find_firewall_directory(self, file_path: Path) -> Optional[Path]:
        current = file_path.parent
        while True:
            if "火墙" in current.name:
                return current
            if current == self.root_path or current.parent == current:
                return None
            current = current.parent

    def scan(
        self,
        progress_callback: Optional[ProgressCallback] = None,
        log_callback: Optional[LogCallback] = None,
        cancel_event: Optional[object] = None,
    ) -> ScanResult:
        root = self.root_path.expanduser()
        summary = ScanSummary(root_path=str(root))
        records_by_group_key: dict[tuple[str, str, str], list[FirewallRecord]] = defaultdict(list)
        seen_by_directory: dict[str, set[tuple[str, str, str]]] = defaultdict(set)
        matched_directories: set[str] = set()

        excel_files = self.collect_excel_files()
        summary.discovered_excel_files = len(excel_files)
        if progress_callback:
            progress_callback(0, len(excel_files), f"发现 {len(excel_files)} 个 Excel 文件")

        for index, file_path in enumerate(excel_files, start=1):
            if cancel_event is not None and getattr(cancel_event, "is_set", lambda: False)():
                summary.cancelled = True
                break

            firewall_dir = self.find_firewall_directory(file_path)
            if firewall_dir is None:
                summary.skipped_outside_firewall_dirs += 1
                if progress_callback:
                    progress_callback(index, len(excel_files), f"已跳过非火墙目录文件：{file_path.name}")
                continue

            matched_directories.add(str(firewall_dir))
            summary.processed_files += 1

            try:
                request_rows, sheet_count = _read_workbook_rows(file_path)
            except Exception as exc:  # pragma: no cover - error path depends on files
                summary.skipped_unreadable_files += 1
                message = f"读取失败：{file_path} - {exc}"
                summary.errors.append(message)
                if log_callback:
                    log_callback(message)
                if progress_callback:
                    progress_callback(index, len(excel_files), f"已跳过无法读取的文件：{file_path.name}")
                continue

            summary.parsed_sheets += sheet_count
            request_dir_key = str(firewall_dir)
            request_dir_rel = self._relative_dir_display(firewall_dir)
            date_text, date_sort = _extract_date_from_request_dir(firewall_dir, self.root_path)

            for request_row in request_rows:
                if not (request_row.source_address and request_row.destination_address and request_row.destination_port):
                    summary.skipped_incomplete_rows += 1
                    continue

                summary.valid_rows += 1
                key = (
                    request_row.source_address,
                    request_row.destination_address,
                    request_row.destination_port,
                )
                if key in seen_by_directory[request_dir_key]:
                    summary.skipped_same_dir_duplicates += 1
                    continue

                seen_by_directory[request_dir_key].add(key)
                record = FirewallRecord(
                    date_text=date_text,
                    date_sort=date_sort,
                    request_dir=request_dir_key,
                    request_dir_rel=request_dir_rel,
                    file_name=file_path.name,
                    file_path=str(file_path),
                    sheet_name=request_row.sheet_name,
                    row_number=request_row.row_number,
                    source_address=request_row.source_address,
                    destination_address=request_row.destination_address,
                    destination_port=request_row.destination_port,
                    normalized_key=key,
                )
                records_by_group_key[key].append(record)

            if progress_callback:
                progress_callback(
                    index,
                    len(excel_files),
                    f"已处理 {index}/{len(excel_files)}：{file_path.name}",
                )

        grouped_items = list(records_by_group_key.items())
        grouped_items.sort(key=lambda item: item[1][0].sort_key())

        groups: list[list[FirewallRecord]] = []
        records: list[FirewallRecord] = []
        duplicate_groups = 0
        duplicate_records = 0

        for group_id, (_, group_records) in enumerate(grouped_items, start=1):
            group_records.sort(key=lambda record: record.sort_key())
            group_size = len(group_records)
            is_duplicate = group_size > 1
            if is_duplicate:
                duplicate_groups += 1
                duplicate_records += group_size
            for record in group_records:
                record.group_id = group_id
                record.group_size = group_size
                record.cross_directory_duplicate = is_duplicate
            groups.append(group_records)

        records = sorted(
            (record for group_records in groups for record in group_records),
            key=lambda record: record.sort_key(),
        )

        summary.matched_directories = len(matched_directories)
        summary.kept_records = len(records)
        summary.cross_directory_duplicate_groups = duplicate_groups
        summary.cross_directory_duplicate_records = duplicate_records

        if log_callback:
            log_callback(f"扫描完成：保留 {summary.kept_records} 条记录，跨目录重复组 {summary.cross_directory_duplicate_groups} 个")

        return ScanResult(summary=summary, records=records, groups=groups)

    def _relative_dir_display(self, request_dir: Path) -> str:
        try:
            relative = request_dir.relative_to(self.root_path)
        except ValueError:
            return request_dir.as_posix()
        text = relative.as_posix()
        return text if text and text != "." else request_dir.name

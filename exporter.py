from __future__ import annotations

from pathlib import Path
from typing import Iterable, Optional

from scanner import FirewallRecord, ScanResult

try:
    from openpyxl import Workbook  # type: ignore
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # type: ignore
    from openpyxl.utils import get_column_letter  # type: ignore
except Exception:  # pragma: no cover - optional dependency handled at runtime
    Workbook = None  # type: ignore
    Alignment = Border = Font = PatternFill = Side = None  # type: ignore
    get_column_letter = None  # type: ignore


DETAIL_HEADERS = [
    "日期",
    "申请目录",
    "文件名",
    "Sheet",
    "来源行",
    "原地址",
    "目的地址",
    "目的端口",
    "重复组",
    "组内数",
    "重复状态",
]


def save_result_to_xlsx(result: ScanResult, output_path: str | Path) -> Path:
    if Workbook is None:
        raise RuntimeError("缺少 openpyxl 依赖，无法导出 Excel 文件。")

    path = Path(output_path)
    if path.suffix.lower() != ".xlsx":
        path = path.with_suffix(".xlsx")
    path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    detail_sheet = workbook.active
    detail_sheet.title = "汇总明细"
    _write_detail_sheet(detail_sheet, result.records)
    _write_duplicate_sheet(workbook.create_sheet("重复组"), result.groups)
    _write_summary_sheet(workbook.create_sheet("统计"), result)

    workbook.save(path)
    return path


def _write_detail_sheet(sheet, records: list[FirewallRecord]) -> None:
    sheet.append(DETAIL_HEADERS)
    _style_header(sheet, 1, len(DETAIL_HEADERS))

    duplicate_fill = PatternFill("solid", fgColor="FFF2CC")
    border = _thin_border()
    alignment = Alignment(vertical="top")
    wrap_alignment = Alignment(vertical="top", wrap_text=True)

    for record in records:
        row_index = sheet.max_row + 1
        sheet.append(
            [
                record.date_text,
                record.request_dir_rel,
                record.file_name,
                record.sheet_name,
                record.row_number,
                record.source_address,
                record.destination_address,
                record.destination_port,
                _group_id(record),
                record.group_size,
                _duplicate_status(record),
            ]
        )

        for column_index in range(1, len(DETAIL_HEADERS) + 1):
            cell = sheet.cell(row=row_index, column=column_index)
            cell.border = border
            cell.alignment = wrap_alignment if column_index in {2, 3, 6, 7} else alignment
            if record.cross_directory_duplicate:
                cell.fill = duplicate_fill

        file_cell = sheet.cell(row=row_index, column=3)
        hyperlink = _file_uri(record.file_path)
        if hyperlink:
            file_cell.hyperlink = hyperlink
            file_cell.font = Font(color="0563C1", underline="single")

    sheet.freeze_panes = "A2"
    if sheet.max_row >= 1:
        sheet.auto_filter.ref = sheet.dimensions
    _auto_width(sheet, max_width=42)


def _write_duplicate_sheet(sheet, groups: list[list[FirewallRecord]]) -> None:
    headers = [
        "重复组",
        "记录数",
        "目录数",
        "日期",
        "原地址",
        "目的地址",
        "目的端口",
        "目录列表",
        "文件列表",
    ]
    sheet.append(headers)
    _style_header(sheet, 1, len(headers))

    duplicate_groups = [group for group in groups if len(group) > 1]
    border = _thin_border()
    alignment = Alignment(vertical="top", wrap_text=True)

    if not duplicate_groups:
        sheet.append(["没有跨目录重复记录"])
    else:
        for group in duplicate_groups:
            first = group[0]
            dates = "；".join(_unique(record.date_text for record in group))
            directories = "；".join(_unique(record.request_dir_rel for record in group))
            files = "；".join(_unique(f"{record.request_dir_rel}/{record.file_name}" for record in group))
            sheet.append(
                [
                    _group_id(first),
                    len(group),
                    len(set(record.request_dir for record in group)),
                    dates,
                    first.source_address,
                    first.destination_address,
                    first.destination_port,
                    directories,
                    files,
                ]
            )

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = border
            cell.alignment = alignment

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    _auto_width(sheet, max_width=55)


def _write_summary_sheet(sheet, result: ScanResult) -> None:
    summary = result.summary
    rows = [
        ("扫描目录", summary.root_path, ""),
        ("扫描状态", "已取消，当前文件为部分结果" if summary.cancelled else "完成", ""),
        ("发现 Excel 文件", summary.discovered_excel_files, "扫描目录下全部 .xlsx/.xlsm/.xls 文件"),
        ("处理 Excel 文件", summary.processed_files, "仅处理位于名称含“火墙”的目录下的文件"),
        ("匹配申请目录", summary.matched_directories, "以最近一级名称含“火墙”的目录作为申请目录"),
        ("读取 Sheet", summary.parsed_sheets, ""),
        ("有效申请行", summary.valid_rows, "C列、F列、H列均有值的行"),
        ("保留记录", summary.kept_records, "同一申请目录内重复项只保留首条"),
        ("同目录去重", summary.skipped_same_dir_duplicates, "同一申请目录内重复的原地址/目的地址/目的端口"),
        ("不完整行", summary.skipped_incomplete_rows, "C列、F列、H列有缺项的行"),
        ("非火墙目录文件", summary.skipped_outside_firewall_dirs, "Excel 文件存在，但不在名称含“火墙”的目录下"),
        ("无法读取文件", summary.skipped_unreadable_files, ""),
        ("跨目录重复组", summary.cross_directory_duplicate_groups, "不同申请目录中出现相同原地址/目的地址/目的端口"),
        ("跨目录重复记录", summary.cross_directory_duplicate_records, ""),
    ]

    sheet.append(["指标", "数值", "说明"])
    _style_header(sheet, 1, 3)
    for row in rows:
        sheet.append(list(row))

    if summary.errors:
        sheet.append([])
        sheet.append(["读取失败明细", "", ""])
        _style_header(sheet, sheet.max_row, 3)
        for error in summary.errors:
            sheet.append([error, "", ""])

    border = _thin_border()
    alignment = Alignment(vertical="top", wrap_text=True)
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, max_col=3):
        for cell in row:
            cell.border = border
            cell.alignment = alignment

    _auto_width(sheet, max_width=60)


def _group_id(record: FirewallRecord) -> str:
    return f"G{record.group_id:04d}"


def _duplicate_status(record: FirewallRecord) -> str:
    return "跨目录重复" if record.cross_directory_duplicate else "唯一"


def _unique(values: Iterable[str]) -> list[str]:
    return list(dict.fromkeys(value for value in values if value))


def _file_uri(path_text: str) -> Optional[str]:
    try:
        path = Path(path_text)
        if not path.is_absolute():
            path = path.resolve()
        return path.as_uri()
    except Exception:
        return None


def _style_header(sheet, row_index: int, column_count: int) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    border = _thin_border()
    for column_index in range(1, column_count + 1):
        cell = sheet.cell(row=row_index, column=column_index)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _thin_border():
    side = Side(style="thin", color="D9E2F3")
    return Border(left=side, right=side, top=side, bottom=side)


def _auto_width(sheet, max_width: int) -> None:
    for column_index, column_cells in enumerate(sheet.iter_cols(), start=1):
        width = 8
        for cell in column_cells:
            if cell.value is None:
                continue
            text = str(cell.value)
            width = max(width, min(max_width, len(text) + 2))
        sheet.column_dimensions[get_column_letter(column_index)].width = width

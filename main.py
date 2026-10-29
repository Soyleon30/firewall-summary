from __future__ import annotations

from datetime import datetime
from pathlib import Path
import argparse
import json
import os
import queue
import subprocess
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import openpyxl
import xlrd

from exporter import save_result_to_xlsx
from scanner import FirewallRecord, FirewallScanner, ScanResult


DEFAULT_ROOT = r"D:\【维护部】"


class FirewallSummaryApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("防火墙申请汇总")
        self.geometry("1320x820")
        self.minsize(1040, 680)

        self.root_path_var = tk.StringVar(value=DEFAULT_ROOT)
        self.status_var = tk.StringVar(value="就绪")
        self.summary_var = tk.StringVar(value=self._summary_text())
        self.progress_value = tk.IntVar(value=0)

        self.event_queue: queue.Queue[tuple] = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.scan_result: ScanResult | None = None
        self.record_by_item: dict[str, FirewallRecord] = {}

        self._build_ui()
        self.after(120, self._poll_events)

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        controls = ttk.Frame(self, padding=(12, 12, 12, 6))
        controls.grid(row=0, column=0, sticky="ew")
        controls.columnconfigure(1, weight=1)

        ttk.Label(controls, text="扫描目录").grid(row=0, column=0, sticky="w")
        ttk.Entry(controls, textvariable=self.root_path_var).grid(row=0, column=1, sticky="ew", padx=(8, 8))
        self.browse_button = ttk.Button(controls, text="选择目录", command=self._browse_directory)
        self.browse_button.grid(row=0, column=2, padx=(0, 8))
        self.scan_button = ttk.Button(controls, text="开始扫描", command=self._start_scan)
        self.scan_button.grid(row=0, column=3, padx=(0, 8))
        self.cancel_button = ttk.Button(controls, text="停止", command=self._cancel_scan, state=tk.DISABLED)
        self.cancel_button.grid(row=0, column=4, padx=(0, 8))
        self.export_button = ttk.Button(controls, text="导出Excel", command=self._export_xlsx, state=tk.DISABLED)
        self.export_button.grid(row=0, column=5)

        self.progress = ttk.Progressbar(controls, variable=self.progress_value, maximum=1)
        self.progress.grid(row=1, column=0, columnspan=6, sticky="ew", pady=(10, 0))

        summary = ttk.Frame(self, padding=(12, 0, 12, 6))
        summary.grid(row=1, column=0, sticky="ew")
        ttk.Label(summary, textvariable=self.summary_var).pack(anchor="w")

        pane = ttk.PanedWindow(self, orient=tk.VERTICAL)
        pane.grid(row=2, column=0, sticky="nsew")

        table_frame = ttk.Frame(pane, padding=(12, 6, 12, 4))
        pane.add(table_frame, weight=4)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = (
            "date",
            "request_dir",
            "file_name",
            "sheet",
            "row",
            "source",
            "destination",
            "port",
            "group",
            "group_size",
            "status",
        )
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=18)
        headings = {
            "date": ("日期", 100),
            "request_dir": ("申请目录", 260),
            "file_name": ("文件名", 200),
            "sheet": ("Sheet", 120),
            "row": ("来源行", 70),
            "source": ("原地址", 220),
            "destination": ("目的地址", 220),
            "port": ("目的端口", 120),
            "group": ("重复组", 90),
            "group_size": ("组内数", 70),
            "status": ("重复状态", 110),
        }
        for column, (title, width) in headings.items():
            self.tree.heading(column, text=title)
            self.tree.column(column, width=width, minwidth=60, anchor=tk.W)
        self.tree.tag_configure("duplicate", background="#fff7dc")
        self.tree.bind("<Double-1>", self._open_selected_source)
        self.tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        log_frame = ttk.Frame(pane, padding=(12, 4, 12, 8))
        pane.add(log_frame, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scroll.set)

        status = ttk.Label(self, textvariable=self.status_var, anchor=tk.W, relief=tk.SUNKEN, padding=(8, 4))
        status.grid(row=3, column=0, sticky="ew")

    def _browse_directory(self) -> None:
        current = Path(self.root_path_var.get().strip('" '))
        initial_dir = current if current.exists() else Path.home()
        selected = filedialog.askdirectory(initialdir=str(initial_dir))
        if selected:
            self.root_path_var.set(selected)

    def _start_scan(self) -> None:
        root_text = self.root_path_var.get().strip().strip('"')
        if not root_text:
            messagebox.showwarning("缺少目录", "请选择要扫描的目录。")
            return

        root_path = Path(root_text)
        if not root_path.exists() or not root_path.is_dir():
            messagebox.showerror("目录无效", f"目录不存在或不可访问：\n{root_text}")
            return

        self._clear_results()
        self.cancel_event.clear()
        self._set_busy(True)
        self.status_var.set("正在扫描")
        self.progress_value.set(0)
        self.progress.configure(maximum=1)
        self._append_log(f"开始扫描：{root_path}")

        self.worker = threading.Thread(target=self._scan_worker, args=(root_path,), daemon=True)
        self.worker.start()

    def _cancel_scan(self) -> None:
        self.cancel_event.set()
        self.status_var.set("正在停止")
        self._append_log("已请求停止扫描，正在等待当前文件处理结束。")

    def _scan_worker(self, root_path: Path) -> None:
        try:
            scanner = FirewallScanner(root_path)
            result = scanner.scan(
                progress_callback=lambda done, total, message: self.event_queue.put(("progress", done, total, message)),
                log_callback=lambda message: self.event_queue.put(("log", message)),
                cancel_event=self.cancel_event,
            )
            self.event_queue.put(("done", result))
        except Exception:
            self.event_queue.put(("error", traceback.format_exc()))

    def _poll_events(self) -> None:
        while True:
            try:
                event = self.event_queue.get_nowait()
            except queue.Empty:
                break

            kind = event[0]
            if kind == "progress":
                _, done, total, message = event
                self._handle_progress(done, total, message)
            elif kind == "log":
                self._append_log(event[1])
            elif kind == "done":
                self._handle_scan_done(event[1])
            elif kind == "error":
                self._handle_scan_error(event[1])

        self.after(120, self._poll_events)

    def _handle_progress(self, done: int, total: int, message: str) -> None:
        if total > 0:
            self.progress.configure(maximum=total)
            self.progress_value.set(done)
        else:
            self.progress.configure(maximum=1)
            self.progress_value.set(0)
        self.status_var.set(message)

    def _handle_scan_done(self, result: ScanResult) -> None:
        self.scan_result = result
        self._populate_table(result.records)
        self._set_busy(False)
        self.summary_var.set(self._summary_text(result))
        status = "扫描已停止，显示部分结果" if result.summary.cancelled else "扫描完成"
        self.status_var.set(status)
        self._append_log(status)

    def _handle_scan_error(self, trace: str) -> None:
        self._set_busy(False)
        self.status_var.set("扫描失败")
        self._append_log(trace)
        messagebox.showerror("扫描失败", trace.splitlines()[-1] if trace else "未知错误")

    def _populate_table(self, records: list[FirewallRecord]) -> None:
        for record in records:
            values = (
                record.date_text,
                record.request_dir_rel,
                record.file_name,
                record.sheet_name,
                record.row_number,
                record.source_address,
                record.destination_address,
                record.destination_port,
                self._group_id(record),
                record.group_size,
                "跨目录重复" if record.cross_directory_duplicate else "唯一",
            )
            tags = ("duplicate",) if record.cross_directory_duplicate else ()
            item_id = self.tree.insert("", tk.END, values=values, tags=tags)
            self.record_by_item[item_id] = record

    def _export_xlsx(self) -> None:
        if self.scan_result is None:
            messagebox.showinfo("无结果", "请先完成扫描。")
            return

        root_path = Path(self.root_path_var.get().strip('" '))
        initial_dir = root_path if root_path.exists() else Path.home()
        default_name = f"防火墙申请汇总_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        selected = filedialog.asksaveasfilename(
            initialdir=str(initial_dir),
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not selected:
            return

        try:
            saved_path = save_result_to_xlsx(self.scan_result, selected)
        except Exception as exc:
            self._append_log(f"导出失败：{exc}")
            messagebox.showerror("导出失败", str(exc))
            return

        self._append_log(f"已导出：{saved_path}")
        messagebox.showinfo("导出完成", f"已保存到：\n{saved_path}")

    def _open_selected_source(self, event: object) -> None:
        item_id = self.tree.focus()
        if not item_id:
            return
        record = self.record_by_item.get(item_id)
        if not record:
            return

        path = Path(record.file_path)
        try:
            if os.name == "nt":
                if path.exists() and path.is_file():
                    subprocess.Popen(["explorer", f"/select,{path}"])
                elif path.exists():
                    os.startfile(path)  # type: ignore[attr-defined]
                else:
                    messagebox.showinfo("文件路径", str(path))
            else:
                messagebox.showinfo("文件路径", str(path))
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))

    def _set_busy(self, busy: bool) -> None:
        state_when_idle = tk.NORMAL if not busy else tk.DISABLED
        self.browse_button.configure(state=state_when_idle)
        self.scan_button.configure(state=state_when_idle)
        self.cancel_button.configure(state=tk.NORMAL if busy else tk.DISABLED)
        self.export_button.configure(state=tk.DISABLED if busy or self.scan_result is None else tk.NORMAL)

    def _clear_results(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.record_by_item.clear()
        self.scan_result = None
        self.summary_var.set(self._summary_text())
        self.export_button.configure(state=tk.DISABLED)

    def _append_log(self, message: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{stamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _summary_text(self, result: ScanResult | None = None) -> str:
        if result is None:
            return "申请目录：0    Excel文件：0    保留记录：0    同目录去重：0    跨目录重复组：0    异常文件：0"
        summary = result.summary
        return (
            f"申请目录：{summary.matched_directories}    "
            f"Excel文件：{summary.processed_files}/{summary.discovered_excel_files}    "
            f"保留记录：{summary.kept_records}    "
            f"同目录去重：{summary.skipped_same_dir_duplicates}    "
            f"跨目录重复组：{summary.cross_directory_duplicate_groups}    "
            f"异常文件：{summary.skipped_unreadable_files}"
        )

    @staticmethod
    def _group_id(record: FirewallRecord) -> str:
        return f"G{record.group_id:04d}" if record.group_id else ""


def run_self_test(output_path: str) -> int:
    result = {
        "ok": False,
        "frozen": bool(getattr(sys, "frozen", False)),
        "python_version": sys.version,
        "python_executable": sys.executable,
        "openpyxl_version": openpyxl.__version__,
        "xlrd_version": xlrd.__version__,
        "tk_version": tk.TkVersion,
    }

    try:
        interpreter = tk.Tcl()
        result["tcl_patchlevel"] = interpreter.eval("info patchlevel")
        result["ok"] = result["frozen"]
    except Exception as exc:
        result["error"] = str(exc)

    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if result["ok"] else 1


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--self-test-output", default="")
    args, _ = parser.parse_known_args()
    return args


def main() -> int:
    args = parse_args()
    if args.self_test:
        if not args.self_test_output:
            return 2
        return run_self_test(args.self_test_output)

    app = FirewallSummaryApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

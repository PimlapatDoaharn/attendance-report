from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
import queue
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, cast

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from attendance_report.processor import ProcessingOptions, find_sheet, process_attendance_report, read_hr_employees

MONTHS = [
    ("January", 1), ("February", 2), ("March", 3), ("April", 4),
    ("May", 5), ("June", 6), ("July", 7), ("August", 8),
    ("September", 9), ("October", 10), ("November", 11), ("December", 12),
]
MONTH_NAME_BY_NUMBER = {number: name for name, number in MONTHS}
MONTH_NUMBER_BY_NAME = {name: number for name, number in MONTHS}

BLUE = "#173B6D"
DARK_BLUE = "#101A27"
LIGHT_BLUE = "#F6F8FB"
BORDER = "#E5EAF2"
TEXT = "#0F172A"
MUTED = "#64748B"
GREEN = "#168447"
BG = "#F8FAFC"
WHITE = "#FFFFFF"


def create_icon(parent: tk.Misc, kind: str, size: int, bg: str, fg: str, border: bool = False) -> tk.Canvas:
    canvas = tk.Canvas(parent, width=size, height=size, bg=bg, highlightthickness=0, bd=0)
    pad = max(size // 7, 4)
    stroke = max(size // 18, 2)
    if border:
        canvas.configure(highlightthickness=1, highlightbackground=BORDER)

    if kind == "chart":
        base = size - pad
        bar_width = max(size // 8, 3)
        for index, height_ratio in enumerate((0.35, 0.6, 0.85)):
            x = pad + index * (bar_width + 4)
            y = base - int((size - 2 * pad) * height_ratio)
            canvas.create_rectangle(x, y, x + bar_width, base, outline=fg, fill="", width=stroke)
    elif kind == "finger":
        center = size / 2
        for offset in range(0, 4):
            radius_x = size * (0.12 + offset * 0.055)
            radius_y = size * (0.20 + offset * 0.07)
            canvas.create_oval(center - radius_x, center - radius_y, center + radius_x, center + radius_y, outline=fg, width=stroke)
        canvas.create_line(center, pad * 1.8, center, center + size * 0.12, fill=fg, width=stroke, capstyle="round")
    elif kind in {"report", "document"}:
        canvas.create_rectangle(pad, pad, size - pad, size - pad, outline=fg, width=stroke)
        canvas.create_line(size - pad * 2, pad, size - pad, pad * 2, fill=fg, width=stroke)
        canvas.create_line(pad * 1.7, size * 0.45, size - pad * 1.7, size * 0.45, fill=fg, width=stroke)
        canvas.create_line(pad * 1.7, size * 0.60, size - pad * 1.7, size * 0.60, fill=fg, width=stroke)
        if kind == "report":
            canvas.create_rectangle(pad * 1.7, size * 0.68, size * 0.42, size - pad * 1.5, outline=fg, width=stroke)
            canvas.create_rectangle(size * 0.52, size * 0.68, size - pad * 1.7, size - pad * 1.5, outline=fg, width=stroke)
    elif kind == "users":
        canvas.create_oval(size * 0.22, size * 0.20, size * 0.45, size * 0.43, outline=fg, width=stroke)
        canvas.create_oval(size * 0.55, size * 0.20, size * 0.78, size * 0.43, outline=fg, width=stroke)
        canvas.create_arc(size * 0.12, size * 0.45, size * 0.55, size * 0.90, start=20, extent=140, outline=fg, width=stroke, style="arc")
        canvas.create_arc(size * 0.45, size * 0.45, size * 0.88, size * 0.90, start=20, extent=140, outline=fg, width=stroke, style="arc")
    elif kind == "calendar":
        canvas.create_rectangle(pad, pad * 1.5, size - pad, size - pad, outline=fg, width=stroke)
        canvas.create_line(pad, size * 0.36, size - pad, size * 0.36, fill=fg, width=stroke)
        for x_ratio in (0.34, 0.5, 0.66):
            for y_ratio in (0.52, 0.68):
                dot = max(size // 28, 2)
                x = size * x_ratio
                y = size * y_ratio
                canvas.create_rectangle(x - dot, y - dot, x + dot, y + dot, outline=fg, fill=fg)
    elif kind == "check":
        canvas.create_oval(pad, pad, size - pad, size - pad, outline=fg, width=stroke)
        cast(Any, canvas).create_line(size * 0.30, size * 0.53, size * 0.45, size * 0.68, size * 0.72, size * 0.35, fill=fg, width=stroke + 1)
    elif kind == "folder":
        canvas.create_line(pad, size * 0.38, size * 0.42, size * 0.38, size * 0.48, size * 0.30, size - pad, size * 0.30, fill=fg, width=stroke)
        canvas.create_rectangle(pad, size * 0.38, size - pad, size - pad, outline=fg, width=stroke)
    elif kind == "bell":
        canvas.create_arc(size * 0.28, size * 0.24, size * 0.72, size * 0.72, start=0, extent=180, outline=fg, width=stroke, style="arc")
        canvas.create_line(size * 0.28, size * 0.48, size * 0.22, size * 0.72, size * 0.78, size * 0.72, size * 0.72, size * 0.48, fill=fg, width=stroke)
        canvas.create_oval(size * 0.46, size * 0.76, size * 0.54, size * 0.84, outline=fg, fill=fg)
    elif kind == "user":
        canvas.create_oval(size * 0.34, size * 0.18, size * 0.66, size * 0.50, outline=fg, width=stroke)
        canvas.create_arc(size * 0.20, size * 0.48, size * 0.80, size * 1.02, start=25, extent=130, outline=fg, width=stroke, style="arc")
    elif kind == "arrow":
        canvas.create_line(pad, size / 2, size - pad, size / 2, fill=fg, width=stroke, capstyle="round")
        canvas.create_line(size - pad * 1.8, size * 0.34, size - pad, size / 2, size - pad * 1.8, size * 0.66, fill=fg, width=stroke, capstyle="round")
    elif kind == "info":
        canvas.create_oval(pad, pad, size - pad, size - pad, outline=fg, width=stroke)
        canvas.create_text(size / 2, size * 0.58, text="i", fill=fg, font=("Arial", max(8, size // 2), "bold"))
    return canvas


def draw_rounded_rectangle(canvas: tk.Canvas, x1: int, y1: int, x2: int, y2: int, radius: int, **kwargs: object) -> int:
    points = [
        x1 + radius,
        y1,
        x2 - radius,
        y1,
        x2,
        y1,
        x2,
        y1 + radius,
        x2,
        y2 - radius,
        x2,
        y2,
        x2 - radius,
        y2,
        x1 + radius,
        y2,
        x1,
        y2,
        x1,
        y2 - radius,
        x1,
        y1 + radius,
        x1,
        y1,
    ]
    return canvas.create_polygon(points, smooth=True, splinesteps=18, **kwargs)


class RoundedFrame(tk.Frame):
    def __init__(
        self,
        parent: tk.Misc,
        *,
        fill: str = WHITE,
        outline: str = BORDER,
        radius: int = 14,
        padding: tuple[int, int, int, int] = (20, 18, 20, 18),
        bg: str = BG,
        min_height: int = 0,
    ) -> None:
        super().__init__(parent, bg=bg)
        self.fill = fill
        self.outline = outline
        self.radius = radius
        self.padding = padding
        self.min_height = min_height
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.canvas.pack(fill="both", expand=True)
        self.body = tk.Frame(self.canvas, bg=fill)
        self.window_id = self.canvas.create_window(padding[0], padding[1], window=self.body, anchor="nw")
        self.canvas.bind("<Configure>", self._redraw)
        self.body.bind("<Configure>", self._sync_height)

    def _sync_height(self, _event: tk.Event) -> None:
        requested = self.body.winfo_reqheight() + self.padding[1] + self.padding[3] + 2
        self.configure(height=max(self.min_height, requested))

    def _redraw(self, event: tk.Event) -> None:
        width = max(event.width, 8)
        height = max(event.height, self.min_height, 8)
        left, top, right, bottom = self.padding
        self.canvas.delete("surface")
        draw_rounded_rectangle(
            self.canvas,
            1,
            1,
            width - 2,
            height - 2,
            self.radius,
            fill=self.fill,
            outline=self.outline,
            width=1,
            tags="surface",
        )
        self.canvas.tag_lower("surface")
        self.canvas.coords(self.window_id, left, top)
        self.canvas.itemconfigure(self.window_id, width=max(width - left - right, 1), height=max(height - top - bottom, 1))


class RoundedButton(tk.Canvas):
    def __init__(
        self,
        parent: tk.Misc,
        text: str,
        command: Callable[[], None],
        *,
        fill: str = WHITE,
        outline: str = "#DCE3EC",
        foreground: str = TEXT,
        radius: int = 10,
        height: int = 48,
    ) -> None:
        super().__init__(parent, height=height, bg=WHITE, highlightthickness=0, bd=0, cursor="hand2")
        self.text = text
        self.command = command
        self.fill = fill
        self.outline = outline
        self.foreground = foreground
        self.radius = radius
        self.bind("<Configure>", self._redraw)
        self.bind("<Button-1>", lambda _event: self.command())
        self.bind("<Enter>", lambda _event: self._set_fill("#F8FAFC"))
        self.bind("<Leave>", lambda _event: self._set_fill(fill))

    def _set_fill(self, fill: str) -> None:
        self.fill = fill
        self._redraw()

    def _redraw(self, _event: tk.Event | None = None) -> None:
        width = max(self.winfo_width(), 12)
        height = max(self.winfo_height(), 12)
        self.delete("all")
        draw_rounded_rectangle(self, 1, 1, width - 2, height - 2, self.radius, fill=self.fill, outline=self.outline, width=1)
        icon_x = 16
        icon_y = (height - 18) // 2
        self.create_rectangle(icon_x, icon_y + 3, icon_x + 18, icon_y + 18, outline=BLUE, width=1.5)
        self.create_line(icon_x, icon_y + 8, icon_x + 18, icon_y + 8, fill=BLUE, width=1.5)
        self.create_line(icon_x + 5, icon_y, icon_x + 5, icon_y + 5, fill=BLUE, width=1.5)
        self.create_line(icon_x + 13, icon_y, icon_x + 13, icon_y + 5, fill=BLUE, width=1.5)
        self.create_text(width / 2 + 8, height / 2, text=self.text, fill=self.foreground, font=("Arial", 10, "bold"))


class SmoothSelect(tk.Canvas):
    def __init__(
        self,
        parent: tk.Misc,
        variable: tk.StringVar | tk.IntVar,
        values: list[str] | list[int],
        *,
        height: int = 48,
        radius: int = 10,
    ) -> None:
        super().__init__(parent, height=height, bg=WHITE, highlightthickness=0, bd=0, cursor="hand2")
        self.variable = variable
        self.radius = radius
        self.menu = tk.Menu(self, tearoff=0, bg=WHITE, fg=TEXT, activebackground=LIGHT_BLUE, activeforeground=TEXT, bd=0)
        for value in values:
            self.menu.add_command(label=str(value), command=lambda selected=value: variable.set(selected))
        self.bind("<Configure>", self._redraw)
        self.bind("<Button-1>", self._open_menu)
        variable.trace_add("write", lambda *_: self._redraw())

    def _open_menu(self, event: tk.Event) -> None:
        self.menu.tk_popup(event.x_root, event.y_root)

    def _redraw(self, _event: tk.Event | None = None) -> None:
        width = max(self.winfo_width(), 80)
        height = max(self.winfo_height(), 48)
        self.delete("all")
        draw_rounded_rectangle(self, 1, 1, width - 2, height - 2, self.radius, fill=WHITE, outline="#DCE3EC", width=1)
        self.create_text(18, height / 2, text=str(self.variable.get()), fill=TEXT, font=("Arial", 11), anchor="w")
        self.create_line(width - 31, height / 2 - 4, width - 24, height / 2 + 3, width - 17, height / 2 - 4, fill=TEXT, width=2.4, capstyle="round", joinstyle="round")


class AttendanceReportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Attendance Report Generator")
        self.geometry("1280x800")
        self.minsize(1120, 720)
        self.configure(bg=BG)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.processing_thread: threading.Thread | None = None
        self.last_output_path: Path | None = None
        self.last_ptvn_output_path: Path | None = None

        today = date.today()
        self.finger_scan_var = tk.StringVar()
        self.report_var = tk.StringVar()
        self.ptvn_report_var = tk.StringVar()
        self.finger_scan_display_var = tk.StringVar(value="ยังไม่ได้เลือกไฟล์")
        self.report_display_var = tk.StringVar(value="ยังไม่ได้เลือกไฟล์")
        self.ptvn_report_display_var = tk.StringVar(value="ยังไม่ได้เลือกไฟล์")
        self.month_var = tk.StringVar(value=MONTH_NAME_BY_NUMBER[today.month])
        self.year_var = tk.IntVar(value=today.year)
        self.working_days_var = tk.IntVar(value=count_weekdays(today.year, today.month))
        self.threshold_var = tk.IntVar(value=80)

        self.records_summary_var = tk.StringVar(value="—")
        self.employees_summary_var = tk.StringVar(value="—")
        self.period_summary_var = tk.StringVar(value=f"{MONTH_NAME_BY_NUMBER[today.month]} {today.year}")
        self.working_summary_var = tk.StringVar(value=f"จำนวนวันที่ทำงาน: {self.working_days_var.get()} วัน")
        self.status_var = tk.StringVar(value="พร้อมใช้งาน")
        self.progress_var = tk.IntVar(value=0)
        self.generated_file_var = tk.StringVar(value="ยังไม่มีไฟล์ที่สร้าง")
        self.generated_time_var = tk.StringVar(value="—")

        self.create_styles()
        self.create_widgets()
        self.bind_traces()
        self.refresh_summary()
        self.after(150, self.drain_log_queue)

    def create_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background=BG)
        style.configure("Card.TFrame", background=WHITE)
        style.configure("Section.TLabel", background=WHITE, foreground=TEXT, font=("Arial", 13, "bold"))
        style.configure("Label.TLabel", background=WHITE, foreground=TEXT, font=("Arial", 10, "bold"))
        style.configure("Muted.TLabel", background=WHITE, foreground=MUTED, font=("Arial", 10))
        style.configure("Success.TLabel", background=WHITE, foreground=GREEN, font=("Arial", 14, "bold"))
        style.configure("Ready.TLabel", background=WHITE, foreground=TEXT, font=("Arial", 13, "bold"))
        style.configure("Primary.TButton", background=DARK_BLUE, foreground=WHITE, font=("Arial", 12, "bold"), padding=15, borderwidth=0)
        style.map("Primary.TButton", background=[("active", "#1E293B"), ("disabled", "#94A3B8")])
        style.configure("Outline.TButton", background=WHITE, foreground=TEXT, font=("Arial", 10, "bold"), padding=10, bordercolor=BORDER)
        style.map("Outline.TButton", background=[("active", LIGHT_BLUE)])
        style.configure("Tool.TButton", background=WHITE, foreground=TEXT, font=("Arial", 11, "bold"), padding=13, bordercolor="#CBD5E1")
        style.map("Tool.TButton", background=[("active", LIGHT_BLUE)])
        style.configure("Horizontal.TProgressbar", background=DARK_BLUE, troughcolor="#EEF2F7", bordercolor="#EEF2F7", lightcolor=DARK_BLUE, darkcolor=DARK_BLUE)
        style.configure("TCombobox", padding=8, fieldbackground=WHITE, background=WHITE, foreground=TEXT)
        style.configure("TSpinbox", padding=8, fieldbackground=WHITE, background=WHITE, foreground=TEXT)

    def create_widgets(self) -> None:
        self.create_header()
        content = ttk.Frame(self, padding=(32, 28, 32, 28))
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=5, uniform="dashboard")
        content.columnconfigure(1, weight=3, uniform="dashboard")
        for row in range(3):
            content.rowconfigure(row, weight=1)

        self.create_input_card(content, row=0, column=0)
        self.create_summary_card(content, row=0, column=1)
        self.create_period_card(content, row=1, column=0)
        self.create_generate_card(content, row=1, column=1)
        self.create_status_card(content, row=2, column=0)
        self.create_file_management_card(content, row=2, column=1)

    def create_header(self) -> None:
        header = tk.Frame(self, bg=WHITE, height=74)
        header.pack(fill="x")
        header.pack_propagate(False)
        create_icon(header, "chart", size=30, bg=WHITE, fg=BLUE).pack(side="left", padx=(34, 18))
        tk.Label(header, text="Attendance Report Generator", bg=WHITE, fg=TEXT, font=("Arial", 17, "bold")).pack(side="left")
        tk.Label(header, text="Admin HR ⌄", bg=WHITE, fg=TEXT, font=("Arial", 11, "bold")).pack(side="right", padx=(10, 34))
        create_icon(header, "user", size=38, bg=LIGHT_BLUE, fg=BLUE).pack(side="right", padx=(0, 10))
        tk.Frame(header, width=1, bg=BORDER).pack(side="right", fill="y", pady=22, padx=18)
        create_icon(header, "bell", size=24, bg=WHITE, fg=TEXT).pack(side="right")
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

    def create_card(
        self,
        parent: ttk.Frame,
        row: int,
        column: int,
        padding: tuple[int, int, int, int] = (24, 22, 24, 22),
        columnspan: int = 1,
    ) -> tk.Frame:
        shell = RoundedFrame(parent, fill=WHITE, outline="#E7ECF3", radius=16, padding=padding, bg=BG, min_height=148)
        shell.grid(row=row, column=column, columnspan=columnspan, sticky="nsew", padx=8, pady=8)
        shell.body.columnconfigure(0, weight=1)
        return shell.body

    def create_input_card(self, parent: ttk.Frame, row: int, column: int) -> None:
        card = self.create_card(parent, row, column)
        ttk.Label(card, text="1. เลือกไฟล์ข้อมูลนำเข้า", style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 18))
        self.add_file_picker(card, 1, "Report Finger Scan (.xlsx)", self.finger_scan_display_var, self.pick_finger_scan)
        self.add_file_picker(card, 2, "Attendance Report (.xlsx)", self.report_display_var, self.pick_report)
        self.add_file_picker(card, 3, "PTVN Report Finger Scan 2026 (.xlsx/.xlsm)", self.ptvn_report_display_var, self.pick_ptvn_report)
        hint = tk.Frame(card, bg=WHITE)
        hint.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        create_icon(hint, "info", size=18, bg=WHITE, fg=MUTED).pack(side="left", padx=(0, 8))
        tk.Label(hint, text="รองรับไฟล์ .xlsx และ PTVN template แบบ .xlsm ขนาดไฟล์ไม่เกิน 50 MB", bg=WHITE, fg=MUTED, font=("Arial", 10)).pack(side="left")

    def add_file_picker(self, parent: tk.Frame, row: int, label: str, display_var: tk.StringVar, command) -> None:
        rounded = RoundedFrame(parent, fill=WHITE, outline="#E6ECF4", radius=11, padding=(18, 13, 18, 13), bg=WHITE, min_height=78)
        rounded.grid(row=row, column=0, sticky="ew", pady=(0, 9))
        shell = rounded.body
        shell.columnconfigure(1, weight=1)
        create_icon(shell, "document", size=34, bg=WHITE, fg=TEXT).grid(row=0, column=0, rowspan=2, padx=(0, 16), pady=4)
        tk.Label(shell, text=label, bg=WHITE, fg=TEXT, font=("Arial", 10, "bold")).grid(row=0, column=1, sticky="sw", pady=(5, 0))
        tk.Label(shell, textvariable=display_var, bg=WHITE, fg=MUTED, font=("Arial", 9)).grid(row=1, column=1, sticky="nw", pady=(2, 5))
        ttk.Button(shell, text="  เลือกไฟล์", style="Outline.TButton", command=command).grid(row=0, column=2, rowspan=2, sticky="e", padx=(12, 0), pady=3)

    def create_period_card(self, parent: ttk.Frame, row: int, column: int) -> None:
        card = self.create_card(parent, row, column)
        ttk.Label(card, text="2. กำหนดช่วงเวลาที่ต้องการประมวลผล", style="Section.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 18))
        self.add_smooth_select(card, 1, 0, "เดือน", self.month_var, [name for name, _ in MONTHS])
        self.add_smooth_select(card, 1, 1, "ปี", self.year_var, list(range(2020, 2101)))
        self.add_smooth_select(card, 1, 2, "จำนวนวันที่ทำงาน (ในเดือน)", self.working_days_var, list(range(1, 32)))
        auto_button = RoundedButton(card, "ตั้งค่าอัตโนมัติ", self.set_auto_working_days, fill=WHITE, outline="#DCE3EC", radius=10, height=48)
        auto_button.grid(row=2, column=3, sticky="ew", pady=(6, 0))
        auto_button.grid_propagate(False)
        hint = tk.Frame(card, bg=WHITE)
        hint.grid(row=3, column=0, columnspan=4, sticky="w", pady=(18, 0))
        create_icon(hint, "info", size=18, bg=WHITE, fg=MUTED).pack(side="left", padx=(0, 8))
        tk.Label(hint, text="ระบบจะยกเว้นวันเสาร์ และวันอาทิตย์ ให้อัตโนมัติ", bg=WHITE, fg=MUTED, font=("Arial", 10)).pack(side="left")
        for col in range(4):
            card.columnconfigure(col, weight=1)

    def add_smooth_select(
        self,
        parent: tk.Frame,
        row: int,
        column: int,
        label: str,
        variable: tk.StringVar | tk.IntVar,
        values: list[str] | list[int],
    ) -> None:
        ttk.Label(parent, text=label, style="Label.TLabel").grid(row=row, column=column, sticky="w", padx=(0, 14))
        select = SmoothSelect(parent, variable, values, height=48, radius=10)
        select.grid(row=row + 1, column=column, sticky="ew", padx=(0, 14), pady=(6, 0))

    def create_generate_card(self, parent: ttk.Frame, row: int, column: int) -> None:
        card = self.create_card(parent, row, column)
        ttk.Label(card, text="3. สร้างรายงาน", style="Section.TLabel").pack(anchor="w", pady=(0, 24))
        self.run_button = ttk.Button(card, text="Generate Report", style="Primary.TButton", command=self.start_processing)
        self.run_button.pack(fill="x", ipady=6)
        tk.Label(card, text="ระบบจะประมวลผลและสร้างไฟล์รายงานโดยอัตโนมัติ", bg=WHITE, fg=MUTED, font=("Arial", 10)).pack(anchor="center", pady=(18, 0))

    def create_status_card(self, parent: ttk.Frame, row: int, column: int) -> None:
        card = self.create_card(parent, row, column)
        ttk.Label(card, text="สถานะการประมวลผล", style="Section.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
        icon_shell = tk.Frame(card, bg=WHITE)
        icon_shell.grid(row=1, column=0, sticky="w", pady=(24, 0))
        create_icon(icon_shell, "check", size=54, bg=WHITE, fg=TEXT).pack(side="left", padx=(0, 18))
        text_shell = tk.Frame(icon_shell, bg=WHITE)
        text_shell.pack(side="left")
        ttk.Label(text_shell, textvariable=self.status_var, style="Ready.TLabel").pack(anchor="w")
        tk.Label(text_shell, text="ระบบพร้อมสำหรับการประมวลผล", bg=WHITE, fg=MUTED, font=("Arial", 9)).pack(anchor="w", pady=(6, 0))
        ttk.Label(card, text="ความคืบหน้า", style="Muted.TLabel").grid(row=1, column=1, sticky="sw", pady=(26, 28), padx=(24, 10))
        ttk.Progressbar(card, variable=self.progress_var, maximum=100, style="Horizontal.TProgressbar").grid(row=1, column=2, sticky="ew", pady=(26, 28), padx=(0, 12))
        self.progress_percent = ttk.Label(card, text="0%", style="Muted.TLabel")
        self.progress_percent.grid(row=1, column=3, sticky="se", pady=(26, 28))
        card.columnconfigure(2, weight=1)

    def create_summary_card(self, parent: ttk.Frame, row: int, column: int) -> None:
        card = self.create_card(parent, row, column)
        ttk.Label(card, text="สรุปข้อมูล", style="Section.TLabel").pack(anchor="w", pady=(0, 18))
        self.add_summary_row(card, "document", "ข้อมูลที่ตรวจพบ (info sheet)", self.records_summary_var)
        self.add_summary_row(card, "users", "พนักงาน (HR)", self.employees_summary_var)
        self.add_summary_row(card, "calendar", "ช่วงเวลาที่ประมวลผล", self.period_summary_var, self.working_summary_var)

    def add_summary_row(self, parent: ttk.Frame, icon_kind: str, label: str, value_var: tk.StringVar, sub_var: tk.StringVar | None = None) -> None:
        row = tk.Frame(parent, bg=WHITE)
        row.pack(fill="x")
        icon_bg = tk.Frame(row, bg=LIGHT_BLUE, width=46, height=46)
        icon_bg.pack_propagate(False)
        icon_bg.pack(side="left", padx=(0, 16), pady=14)
        create_icon(icon_bg, icon_kind, size=28, bg=LIGHT_BLUE, fg=BLUE).pack(expand=True)
        text_frame = tk.Frame(row, bg=WHITE)
        text_frame.pack(side="left", fill="x", expand=True, pady=10)
        tk.Label(text_frame, text=label, bg=WHITE, fg=TEXT, font=("Arial", 10)).pack(anchor="w")
        tk.Label(text_frame, textvariable=value_var, bg=WHITE, fg=TEXT, font=("Arial", 18, "bold")).pack(anchor="w")
        if sub_var:
            tk.Label(text_frame, textvariable=sub_var, bg=WHITE, fg=MUTED, font=("Arial", 10)).pack(anchor="w", pady=(4, 0))
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x")

    def create_file_management_card(self, parent: ttk.Frame, row: int, column: int) -> None:
        card = self.create_card(parent, row, column)
        ttk.Label(card, text="จัดการไฟล์", style="Section.TLabel").pack(anchor="w", pady=(0, 20))
        button_row = tk.Frame(card, bg=WHITE)
        button_row.pack(fill="x")
        button_row.columnconfigure(0, weight=1)
        button_row.columnconfigure(1, weight=1)
        ttk.Button(button_row, text="  เปิดไฟล์", style="Tool.TButton", command=self.open_output_file).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(button_row, text="  เปิดโฟลเดอร์ปลายทาง", style="Tool.TButton", command=self.open_output_folder).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        tk.Label(card, textvariable=self.generated_file_var, bg=WHITE, fg=MUTED, font=("Arial", 9), wraplength=420, justify="left").pack(anchor="w", pady=(18, 4))
        tk.Label(card, textvariable=self.generated_time_var, bg=WHITE, fg=MUTED, font=("Arial", 9)).pack(anchor="w")

    def bind_traces(self) -> None:
        self.month_var.trace_add("write", lambda *_: self.on_period_changed())
        self.year_var.trace_add("write", lambda *_: self.on_period_changed())
        self.working_days_var.trace_add("write", lambda *_: self.refresh_summary())

    def pick_finger_scan(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            self.finger_scan_var.set(path)
            self.refresh_summary()

    def pick_report(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            self.report_var.set(path)
            self.refresh_summary()

    def pick_ptvn_report(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if path:
            self.ptvn_report_var.set(path)
            self.refresh_summary()

    def on_period_changed(self) -> None:
        try:
            self.set_auto_working_days()
        except Exception:
            self.refresh_summary()

    def set_auto_working_days(self) -> None:
        self.working_days_var.set(count_weekdays(int(self.year_var.get()), self.selected_month()))
        self.refresh_summary()

    def selected_month(self) -> int:
        return MONTH_NUMBER_BY_NAME.get(self.month_var.get(), date.today().month)

    def refresh_summary(self) -> None:
        self.period_summary_var.set(f"{self.month_var.get()} {self.year_var.get()}")
        self.working_summary_var.set(f"จำนวนวันที่ทำงาน: {self.working_days_var.get()} วัน")
        self.finger_scan_display_var.set(self.display_file_name(self.finger_scan_var.get()))
        self.report_display_var.set(self.display_file_name(self.report_var.get()))
        self.ptvn_report_display_var.set(self.display_file_name(self.ptvn_report_var.get()))
        self.records_summary_var.set(self.detect_finger_records(self.finger_scan_var.get()))
        self.employees_summary_var.set(self.detect_hr_employees(self.report_var.get()))

    def display_file_name(self, path: str) -> str:
        return Path(path).name if path else "ยังไม่ได้เลือกไฟล์"

    def detect_finger_records(self, path: str) -> str:
        if not path:
            return "—"
        try:
            workbook = load_workbook(path, read_only=True, data_only=True)
            sheet = cast(Worksheet, workbook["info (used)"] if "info (used)" in workbook.sheetnames else workbook["info"] if "info" in workbook.sheetnames else workbook.active)
            records = max((sheet.max_row or 0) - 1, 0)
            workbook.close()
            return f"{records:,}"
        except Exception:
            return "Unable to read"

    def detect_hr_employees(self, path: str) -> str:
        if not path:
            return "—"
        try:
            workbook = load_workbook(path, read_only=False, data_only=False)
            sheet = find_sheet(workbook, "ฐานข้อมูลพนักงาน (HR)")
            count = len(read_hr_employees(sheet)) if sheet else 0
            workbook.close()
            return f"{count:,}"
        except Exception:
            return "Unable to read"

    def start_processing(self) -> None:
        if self.processing_thread and self.processing_thread.is_alive():
            return
        missing = [
            label
            for label, value in [
                ("Report Finger Scan", self.finger_scan_var.get()),
                ("Attendance Report", self.report_var.get()),
                ("PTVN Report Finger Scan 2026", self.ptvn_report_var.get()),
            ]
            if not value
        ]
        if missing:
            messagebox.showwarning("Missing input", "Please select: " + ", ".join(missing))
            return

        self.progress_var.set(5)
        self.progress_percent.configure(text="5%")
        self.status_var.set("กำลังประมวลผล...")
        self.run_button.configure(state="disabled")
        self.processing_thread = threading.Thread(target=self.process_worker, daemon=True)
        self.processing_thread.start()

    def process_worker(self) -> None:
        try:
            options = ProcessingOptions(
                month=self.selected_month(),
                year=int(self.year_var.get()),
                fuzzy_threshold=int(self.threshold_var.get()),
                working_days=int(self.working_days_var.get()),
            )
            output = self.default_output_path()
            ptvn_output = self.default_ptvn_output_path()
            result = process_attendance_report(
                self.finger_scan_var.get(),
                self.report_var.get(),
                output,
                options,
                ptvn_report_path=self.ptvn_report_var.get(),
                ptvn_output_path=ptvn_output,
                log=self.log_queue.put,
            )
            self.log_queue.put(f"DONE::{result}::{ptvn_output}")
        except Exception as exc:
            self.log_queue.put(f"ERROR::{exc}")

    def default_output_path(self) -> Path:
        report_path = Path(self.report_var.get())
        return report_path.with_name(f"Attendance_Report_{self.month_var.get()}{int(self.year_var.get())}.xlsx")

    def default_ptvn_output_path(self) -> Path:
        ptvn_path = Path(self.ptvn_report_var.get())
        output_suffix = ".xlsm" if ptvn_path.suffix.lower() == ".xlsm" else ".xlsx"
        return ptvn_path.with_name(f"PTVN_Report_Finger_Scan_{self.month_var.get()}{int(self.year_var.get())}{output_suffix}")

    def drain_log_queue(self) -> None:
        try:
            while True:
                message = self.log_queue.get_nowait()
                if message.startswith("DONE::"):
                    output_text, ptvn_output_text = message.replace("DONE::", "", 1).split("::", 1)
                    output = Path(output_text)
                    ptvn_output = Path(ptvn_output_text)
                    self.last_output_path = output
                    self.last_ptvn_output_path = ptvn_output
                    self.progress_var.set(100)
                    self.progress_percent.configure(text="100%")
                    self.status_var.set("เสร็จสมบูรณ์")
                    self.generated_file_var.set(f"1. {output.name}\n2. {ptvn_output.name}")
                    self.generated_time_var.set(datetime.now().strftime("%d %b %Y %H:%M:%S"))
                    self.run_button.configure(state="normal")
                    self.refresh_summary()
                elif message.startswith("ERROR::"):
                    error = message.replace("ERROR::", "", 1)
                    self.progress_var.set(0)
                    self.progress_percent.configure(text="0%")
                    self.status_var.set("เกิดข้อผิดพลาด")
                    self.run_button.configure(state="normal")
                    messagebox.showerror("Error", error)
                else:
                    self.advance_progress(message)
        except queue.Empty:
            pass
        self.after(150, self.drain_log_queue)

    def advance_progress(self, message: str) -> None:
        next_value = min(max(self.progress_var.get() + 12, 12), 92)
        self.progress_var.set(next_value)
        self.progress_percent.configure(text=f"{next_value}%")
        self.status_var.set(message)

    def open_output_file(self) -> None:
        output_paths = [path for path in [self.last_output_path, self.last_ptvn_output_path] if path and path.exists()]
        if output_paths:
            for output_path in output_paths:
                subprocess.run(["open", str(output_path)], check=False)
        else:
            messagebox.showinfo("No file", "No generated files are available yet.")

    def open_output_folder(self) -> None:
        if self.last_output_path and self.last_output_path.exists():
            subprocess.run(["open", str(self.last_output_path.parent)], check=False)
        elif self.last_ptvn_output_path and self.last_ptvn_output_path.exists():
            subprocess.run(["open", str(self.last_ptvn_output_path.parent)], check=False)
        else:
            messagebox.showinfo("No folder", "No generated file is available yet.")


def count_weekdays(year: int, month: int) -> int:
    current = date(year, month, 1)
    end = date(year + 1, 1, 1) if month == 12 else date(year, month + 1, 1)
    total = 0
    while current < end:
        if current.weekday() < 5:
            total += 1
        current += timedelta(days=1)
    return total


def main() -> None:
    app = AttendanceReportApp()
    app.mainloop()


if __name__ == "__main__":
    main()

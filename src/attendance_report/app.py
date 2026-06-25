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

from attendance_report.processor import (
    ProcessingOptions,
    find_sheet,
    process_attendance_report,
    read_hr_employees,
)

# ── Theme ──────────────────────────────────────────────────────────
BG            = "#F2F5FB"
CARD          = "#FFFFFF"
BORDER        = "#E2E8F0"
BORDER_LIGHT  = "#EBF0F7"
PRIMARY       = "#3B5BDB"
PRIMARY_FG    = "#FFFFFF"
PRIMARY_LIGHT = "#EEF2FF"
TEXT          = "#1A1D2E"
MUTED_BG      = "#EBF0F7"
MUTED_FG      = "#6B7280"
EMERALD       = "#059669"
EMERALD_BG    = "#ECFDF5"
EMERALD_DARK  = "#047857"
AMBER         = "#D97706"
AMBER_BG      = "#FFFBEB"

MONTHS = [
    ("January", 1), ("February", 2), ("March", 3), ("April", 4),
    ("May", 5), ("June", 6), ("July", 7), ("August", 8),
    ("September", 9), ("October", 10), ("November", 11), ("December", 12),
]
MONTH_NAME_BY_NUMBER = {n: nm for nm, n in MONTHS}
MONTH_NUMBER_BY_NAME = {nm: n for nm, n in MONTHS}

FILE_CONFIGS = [
    ("finger", "Report Finger Scan",  "Biometric scan export",    "DRAFT"),
    ("draft",  "Draft Attendance",    "Pre-processed draft file", "DRAFT"),
    ("ptvn",   "Attendance Report",   "Final attendance sheet",   "FINAL"),
]


def count_weekdays(year: int, month: int) -> int:
    cur = date(year, month, 1)
    end = date(year + 1, 1, 1) if month == 12 else date(year, month + 1, 1)
    n = 0
    while cur < end:
        if cur.weekday() < 5:
            n += 1
        cur += timedelta(days=1)
    return n


def draw_rounded_rect(canvas: tk.Canvas, x1: int, y1: int, x2: int, y2: int,
                      r: int, **kw: object) -> int:
    pts = [
        x1+r, y1,  x2-r, y1,  x2, y1,    x2, y1+r,
        x2, y2-r,  x2, y2,    x2-r, y2,  x1+r, y2,
        x1, y2,    x1, y2-r,  x1, y1+r,  x1, y1,
    ]
    return canvas.create_polygon(pts, smooth=True, splinesteps=18, **kw)


# ── Widgets ────────────────────────────────────────────────────────

class RCard(tk.Frame):
    def __init__(self, parent: tk.Misc, *,
                 padding: tuple[int, int, int, int] = (20, 18, 20, 18),
                 radius: int = 14, outer_bg: str = BG) -> None:
        super().__init__(parent, bg=outer_bg)
        self._r = radius
        self._pad = padding
        self._cv = tk.Canvas(self, bg=outer_bg, highlightthickness=0, bd=0)
        self._cv.pack(fill="both", expand=True)
        self.body = tk.Frame(self._cv, bg=CARD)
        self._win = self._cv.create_window(padding[0], padding[1], window=self.body, anchor="nw")
        self._cv.bind("<Configure>", self._redraw)
        self.body.bind("<Configure>", self._sync)

    def _sync(self, _: tk.Event) -> None:
        h = self.body.winfo_reqheight() + self._pad[1] + self._pad[3]
        self.configure(height=max(h, 40))

    def _redraw(self, e: tk.Event) -> None:
        w, h = max(e.width, 8), max(e.height, 8)
        self._cv.delete("bg")
        draw_rounded_rect(self._cv, 1, 1, w-2, h-2, self._r,
                          fill=CARD, outline=BORDER, width=1, tags="bg")
        self._cv.tag_lower("bg")
        l, t, r, b = self._pad
        self._cv.coords(self._win, l, t)
        self._cv.itemconfigure(self._win, width=max(w-l-r, 1), height=max(h-t-b, 1))


class FileCard(tk.Frame):
    def __init__(self, parent: tk.Misc, label: str, desc: str, badge: str,
                 command: Callable[[], None]) -> None:
        super().__init__(parent, bg=BG)
        self._label_text = label
        self._desc_text  = desc
        self._badge      = badge
        self._command    = command
        self._filename: str | None = None
        self._cv   = tk.Canvas(self, bg=BG, highlightthickness=0, bd=0, cursor="hand2")
        self._cv.pack(fill="both", expand=True)
        self.body  = tk.Frame(self._cv, bg=CARD)
        self._win  = self._cv.create_window(14, 14, window=self.body, anchor="nw")
        self._cv.bind("<Configure>", self._redraw)
        self.body.bind("<Configure>", self._sync)
        self._cv.bind("<Button-1>", lambda _: command())
        self._rebuild()

    def _rebuild(self) -> None:
        for w in self.body.winfo_children():
            w.destroy()
        ready = self._filename is not None
        bg = EMERALD_BG if ready else CARD
        self.body.configure(bg=bg)
        badge_bg = PRIMARY_LIGHT if self._badge == "FINAL" else MUTED_BG
        badge_fg = PRIMARY       if self._badge == "FINAL" else MUTED_FG
        tk.Label(self.body, text=self._badge, bg=badge_bg, fg=badge_fg,
                 font=("Arial", 8, "bold"), padx=5, pady=1).pack(anchor="e", pady=(0, 6))
        tk.Label(self.body, text=self._label_text, bg=bg, fg=TEXT,
                 font=("Arial", 10, "bold"), anchor="w").pack(fill="x")
        tk.Label(self.body, text=self._desc_text, bg=bg, fg=MUTED_FG,
                 font=("Arial", 9), anchor="w").pack(fill="x", pady=(2, 8))
        if ready:
            info = tk.Frame(self.body, bg="#D1FAE5")
            info.pack(fill="x")
            tk.Label(info, text=f"✓  {self._filename}", bg="#D1FAE5", fg=EMERALD_DARK,
                     font=("Arial", 9, "bold"), anchor="w", wraplength=170,
                     justify="left").pack(fill="x", padx=8, pady=5)
        else:
            tk.Label(self.body, text="↑  Click to upload .xlsx", bg=bg, fg=MUTED_FG,
                     font=("Arial", 9), anchor="w").pack(fill="x")

    def set_file(self, name: str | None) -> None:
        self._filename = name
        self._rebuild()
        self.body.update_idletasks()
        self.configure(height=max(self.body.winfo_reqheight() + 28, 140))

    def _sync(self, _: tk.Event) -> None:
        self.configure(height=max(self.body.winfo_reqheight() + 28, 140))

    def _redraw(self, e: tk.Event) -> None:
        w, h = max(e.width, 8), max(e.height, 8)
        self._cv.delete("bg")
        ready = self._filename is not None
        draw_rounded_rect(self._cv, 1, 1, w-2, h-2, 12,
                          fill=(EMERALD_BG if ready else CARD),
                          outline=(EMERALD if ready else BORDER),
                          width=2, dash=(5, 4), tags="bg")
        self._cv.tag_lower("bg")
        self._cv.coords(self._win, 14, 14)
        self._cv.itemconfigure(self._win, width=max(w-28, 1), height=max(h-28, 1))


class SummaryRow(tk.Frame):
    def __init__(self, parent: tk.Misc, icon: str, label: str,
                 value_var: tk.StringVar,
                 icon_bg: str = PRIMARY_LIGHT, icon_fg: str = PRIMARY) -> None:
        super().__init__(parent, bg=MUTED_BG,
                         highlightthickness=1, highlightbackground=BORDER_LIGHT)
        inner = tk.Frame(self, bg=MUTED_BG)
        inner.pack(fill="both", expand=True, padx=10, pady=9)
        ib = tk.Frame(inner, bg=icon_bg, width=32, height=32)
        ib.pack_propagate(False)
        ib.pack(side="left", padx=(0, 10))
        tk.Label(ib, text=icon, bg=icon_bg, fg=icon_fg, font=("Arial", 14)).pack(expand=True)
        txt = tk.Frame(inner, bg=MUTED_BG)
        txt.pack(side="left", fill="x", expand=True)
        tk.Label(txt, text=label, bg=MUTED_BG, fg=MUTED_FG, font=("Arial", 9)).pack(anchor="w")
        tk.Label(txt, textvariable=value_var, bg=MUTED_BG, fg=TEXT,
                 font=("Arial", 11, "bold")).pack(anchor="w")


class CheckItem(tk.Frame):
    def __init__(self, parent: tk.Misc, label: str, sub: str, done: bool = False) -> None:
        super().__init__(parent, bg=CARD)
        self._ind = tk.Canvas(self, width=18, height=18, bg=CARD, highlightthickness=0)
        self._ind.pack(side="left", anchor="n", padx=(0, 10), pady=2)
        txt = tk.Frame(self, bg=CARD)
        txt.pack(side="left", fill="x", expand=True)
        self._lbl = tk.Label(txt, text=label, bg=CARD, font=("Arial", 10))
        self._lbl.pack(anchor="w")
        self._sub = tk.Label(txt, text=sub, bg=CARD, fg=MUTED_FG, font=("Arial", 9))
        self._sub.pack(anchor="w")
        self._draw(done, False)

    def _draw(self, done: bool, active: bool) -> None:
        self._ind.delete("all")
        if done:
            self._ind.create_oval(1, 1, 16, 16, fill=EMERALD_BG, outline=EMERALD, width=1.5)
            self._ind.create_text(9, 9, text="✓", fill=EMERALD, font=("Arial", 8, "bold"))
        elif active:
            self._ind.create_oval(1, 1, 16, 16, fill=PRIMARY_LIGHT, outline=PRIMARY, width=1.5)
            self._ind.create_text(9, 9, text="▶", fill=PRIMARY, font=("Arial", 7))
        else:
            self._ind.create_oval(1, 1, 16, 16, fill=CARD, outline=BORDER, width=1.5)

    def update(self, label: str, sub: str, done: bool, active: bool = False) -> None:
        self._draw(done, active)
        self._lbl.configure(text=label, fg=TEXT if done else (PRIMARY if active else MUTED_FG),
                             font=("Arial", 10, "bold" if done else "normal"))
        self._sub.configure(text=sub)


# ── App ────────────────────────────────────────────────────────────

class AttendanceReportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Attendance Report")
        self.geometry("1280x820")
        self.minsize(1100, 700)
        self.configure(bg=BG)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.processing_thread: threading.Thread | None = None
        self.last_output_path: Path | None = None
        self.last_ptvn_output_path: Path | None = None

        today = date.today()
        self.finger_scan_var  = tk.StringVar()
        self.report_var       = tk.StringVar()
        self.ptvn_report_var  = tk.StringVar()
        self.month_var        = tk.StringVar(value=MONTH_NAME_BY_NUMBER[today.month])
        self.year_var         = tk.IntVar(value=today.year)
        self.working_days_var = tk.IntVar(value=count_weekdays(today.year, today.month))
        self._sum_period      = tk.StringVar()
        self._sum_wdays       = tk.StringVar()
        self._sum_files       = tk.StringVar()

        self._is_generating = False
        self._is_generated  = False
        self._progress      = 0.0
        self._gen_time      = ""

        self._build_ui()
        self._bind_traces()
        self._refresh()
        self.after(150, self._drain_queue)

    # ── Build UI ──────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self._build_header()
        self._build_main()
        self._build_footer()

    def _build_header(self) -> None:
        header = tk.Frame(self, bg=CARD, height=56)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        logo = tk.Canvas(header, width=32, height=32, bg=CARD, highlightthickness=0)
        logo.pack(side="left", padx=(28, 10), pady=12)
        draw_rounded_rect(logo, 0, 0, 32, 32, 8, fill=PRIMARY, outline="")
        logo.create_text(16, 16, text="≡", fill=PRIMARY_FG, font=("Arial", 14, "bold"))

        title_f = tk.Frame(header, bg=CARD)
        title_f.pack(side="left")
        tk.Label(title_f, text="Attendance Report", bg=CARD, fg=TEXT,
                 font=("Arial", 12, "bold")).pack(anchor="w")
        tk.Label(title_f, text="HR Reporting System", bg=CARD, fg=MUTED_FG,
                 font=("Arial", 9)).pack(anchor="w")

        right = tk.Frame(header, bg=CARD)
        right.pack(side="right", padx=28)
        self._status_badge = tk.Label(right, text="● Idle", bg=MUTED_BG, fg=MUTED_FG,
                                       font=("Arial", 10, "bold"), padx=10, pady=3)
        self._status_badge.pack(side="right")
        tk.Frame(right, bg=BORDER, width=1, height=20).pack(side="right", padx=10)
        self._header_period = tk.Label(right, bg=CARD, fg=MUTED_FG, font=("Arial", 10))
        self._header_period.pack(side="right")

    def _build_main(self) -> None:
        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True, padx=28, pady=20)
        outer.columnconfigure(0, weight=1)
        outer.columnconfigure(1, minsize=310)
        outer.rowconfigure(0, weight=1)

        left = tk.Frame(outer, bg=BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        self._build_upload_section(left)
        self._build_period_section(left)
        self._build_generate_section(left)
        self._build_file_mgmt_section(left)

        right = tk.Frame(outer, bg=BG)
        right.grid(row=0, column=1, sticky="new")
        self._build_summary_panel(right)
        self._build_checklist_panel(right)

    def _step_header(self, parent: tk.Frame, step: str, title: str,
                     right_var: tk.StringVar | None = None) -> None:
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill="x", pady=(0, 14))
        circ = tk.Canvas(row, width=24, height=24, bg=CARD, highlightthickness=0)
        circ.pack(side="left", padx=(0, 10))
        draw_rounded_rect(circ, 0, 0, 24, 24, 12, fill=PRIMARY, outline="")
        circ.create_text(12, 12, text=step, fill=PRIMARY_FG, font=("Arial", 10, "bold"))
        tk.Label(row, text=title, bg=CARD, fg=TEXT, font=("Arial", 12, "bold")).pack(side="left")
        if right_var is not None:
            tk.Label(row, textvariable=right_var, bg=CARD, fg=MUTED_FG,
                     font=("Arial", 10)).pack(side="right")

    def _build_upload_section(self, parent: tk.Frame) -> None:
        card = RCard(parent, padding=(20, 18, 20, 18))
        card.pack(fill="x", pady=(0, 10))
        self._upload_counter = tk.StringVar(value="0/3 uploaded")
        self._step_header(card.body, "1", "Upload Attendance Files", self._upload_counter)

        grid = tk.Frame(card.body, bg=CARD)
        grid.pack(fill="x")
        for c in range(3):
            grid.columnconfigure(c, weight=1, uniform="fc")

        self._file_cards: dict[str, FileCard] = {}
        pick_map = {"finger": self._pick_finger, "draft": self._pick_report, "ptvn": self._pick_ptvn}
        for col, (key, label, desc, badge) in enumerate(FILE_CONFIGS):
            fc = FileCard(grid, label, desc, badge, pick_map[key])
            fc.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 6, 0))
            self._file_cards[key] = fc

    def _build_period_section(self, parent: tk.Frame) -> None:
        card = RCard(parent, padding=(20, 18, 20, 18))
        card.pack(fill="x", pady=(0, 10))
        self._step_header(card.body, "2", "Select Report Period")

        grid = tk.Frame(card.body, bg=CARD)
        grid.pack(fill="x")
        for c in range(3):
            grid.columnconfigure(c, weight=1, uniform="ps")

        for col, (label, var, vals) in enumerate([
            ("MONTH",        self.month_var,        [nm for nm, _ in MONTHS]),
            ("YEAR",         self.year_var,          list(range(2020, 2101))),
            ("WORKING DAYS", self.working_days_var,  list(range(1, 32))),
        ]):
            cf = tk.Frame(grid, bg=CARD)
            cf.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 10, 0))
            tk.Label(cf, text=label, bg=CARD, fg=MUTED_FG,
                     font=("Arial", 9, "bold")).pack(anchor="w", pady=(0, 4))
            cb = ttk.Combobox(cf, textvariable=var, values=[str(v) for v in vals],
                               font=("Arial", 11), state="readonly")
            cb.pack(fill="x", ipady=4)

    def _build_generate_section(self, parent: tk.Frame) -> None:
        card = RCard(parent, padding=(20, 18, 20, 18))
        card.pack(fill="x", pady=(0, 10))
        b = card.body

        top = tk.Frame(b, bg=CARD)
        top.pack(fill="x", pady=(0, 12))
        txt = tk.Frame(top, bg=CARD)
        txt.pack(side="left", fill="x", expand=True)
        tk.Label(txt, text="Generate Report", bg=CARD, fg=TEXT,
                 font=("Arial", 12, "bold")).pack(anchor="w")
        self._gen_subtitle = tk.Label(txt, text="Upload at least one file to continue",
                                       bg=CARD, fg=MUTED_FG, font=("Arial", 10))
        self._gen_subtitle.pack(anchor="w", pady=(2, 0))

        self._gen_btn = tk.Canvas(top, width=140, height=38, bg=CARD,
                                   highlightthickness=0, cursor="hand2")
        self._gen_btn.pack(side="right", padx=(14, 0))
        self._gen_btn.bind("<Button-1>", lambda _: self._start_processing())
        self._gen_btn.bind("<Enter>",    lambda _: self._gen_hover(True))
        self._gen_btn.bind("<Leave>",    lambda _: self._gen_hover(False))
        self._draw_gen_btn()

        status_row = tk.Frame(b, bg=CARD)
        status_row.pack(fill="x", pady=(0, 6))
        self._status_lbl = tk.Label(status_row, text="Awaiting input", bg=CARD,
                                     fg=MUTED_FG, font=("Courier", 10))
        self._status_lbl.pack(side="left")
        self._pct_lbl = tk.Label(status_row, text="0%", bg=CARD, fg=TEXT,
                                  font=("Courier", 10, "bold"))
        self._pct_lbl.pack(side="right")

        prog_bg = tk.Frame(b, bg=MUTED_BG, height=8)
        prog_bg.pack(fill="x")
        prog_bg.pack_propagate(False)
        self._prog_fill = tk.Frame(prog_bg, bg=PRIMARY, height=8)
        self._prog_fill.place(x=0, y=0, relheight=1.0, relwidth=0.0)

        self._output_frame = tk.Frame(b, bg=CARD)
        self._output_frame.pack(fill="x", pady=(12, 0))

    def _draw_gen_btn(self, hover: bool = False, disabled: bool = False) -> None:
        self._gen_btn.delete("all")
        if disabled:
            fill, fg = MUTED_BG, MUTED_FG
        elif hover:
            fill, fg = "#2F4CC4", PRIMARY_FG
        else:
            fill, fg = PRIMARY, PRIMARY_FG
        draw_rounded_rect(self._gen_btn, 0, 0, 140, 38, 10, fill=fill, outline="")
        txt = "Generating…" if self._is_generating else "▶  Generate"
        self._gen_btn.create_text(70, 19, text=txt, fill=fg, font=("Arial", 10, "bold"))

    def _gen_hover(self, on: bool) -> None:
        if not self._is_generating:
            self._draw_gen_btn(hover=on)

    def _build_file_mgmt_section(self, parent: tk.Frame) -> None:
        card = RCard(parent, padding=(18, 14, 18, 14))
        card.pack(fill="x")
        b = card.body
        left_txt = tk.Frame(b, bg=CARD)
        left_txt.pack(side="left", fill="x", expand=True)
        tk.Label(left_txt, text="File Management", bg=CARD, fg=TEXT,
                 font=("Arial", 11, "bold")).pack(anchor="w")
        tk.Label(left_txt, text="Access uploaded sources and generated outputs",
                 bg=CARD, fg=MUTED_FG, font=("Arial", 9)).pack(anchor="w")
        btn_row = tk.Frame(b, bg=CARD)
        btn_row.pack(side="right")
        self._outline_btn(btn_row, "🗂  Open Uploaded Files", self._open_folder).pack(side="left", padx=(0, 8))
        self._outline_btn(btn_row, "↓  Open All Generated",  self._open_file).pack(side="left")

    def _outline_btn(self, parent: tk.Frame, text: str, cmd: Callable[[], None]) -> tk.Canvas:
        c = tk.Canvas(parent, width=185, height=36, bg=CARD, highlightthickness=0, cursor="hand2")
        c.bind("<Button-1>", lambda _: cmd())
        c.bind("<Configure>", lambda e, c=c, t=text: self._redraw_outline(c, t))
        self._redraw_outline(c, text)
        return c

    def _redraw_outline(self, c: tk.Canvas, text: str) -> None:
        c.delete("all")
        w, h = max(c.winfo_width(), 185), max(c.winfo_height(), 36)
        draw_rounded_rect(c, 0, 0, w-1, h-1, 10, fill=CARD, outline=BORDER, width=1)
        c.create_text(w//2, h//2, text=text, fill=TEXT, font=("Arial", 10, "bold"))

    def _build_summary_panel(self, parent: tk.Frame) -> None:
        card = RCard(parent, padding=(18, 16, 18, 16))
        card.pack(fill="x", pady=(0, 10))
        b = card.body
        tk.Label(b, text="SUMMARY", bg=CARD, fg=MUTED_FG,
                 font=("Arial", 9, "bold")).pack(anchor="w", pady=(0, 12))
        SummaryRow(b, "📅", "Period",        self._sum_period, PRIMARY_LIGHT, PRIMARY).pack(fill="x", pady=(0, 6))
        SummaryRow(b, "⏱",  "Working Days",  self._sum_wdays,  AMBER_BG,     AMBER).pack(fill="x", pady=(0, 6))
        SummaryRow(b, "📊", "Uploaded Files", self._sum_files,  EMERALD_BG,   EMERALD).pack(fill="x")

    def _build_checklist_panel(self, parent: tk.Frame) -> None:
        card = RCard(parent, padding=(18, 16, 18, 16))
        card.pack(fill="x")
        b = card.body
        tk.Label(b, text="PROCESS CHECKLIST", bg=CARD, fg=MUTED_FG,
                 font=("Arial", 9, "bold")).pack(anchor="w", pady=(0, 12))
        self._check_items = [
            CheckItem(b, "Upload source files",        "No files uploaded"),
            CheckItem(b, "Configure report period",    ""),
            CheckItem(b, "Set working days",           ""),
            CheckItem(b, "Generate attendance report", "Pending"),
        ]
        for item in self._check_items:
            item.pack(fill="x", pady=(0, 10))

    def _build_footer(self) -> None:
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")
        footer = tk.Frame(self, bg=CARD, height=40)
        footer.pack(fill="x")
        footer.pack_propagate(False)
        tk.Label(footer,
                 text="🔒  All files are processed locally on this machine. No employee data is transmitted to external servers.",
                 bg=CARD, fg=MUTED_FG, font=("Arial", 9)).pack(side="left", padx=28, anchor="w")

    # ── Refresh ───────────────────────────────────────────────────

    def _refresh(self) -> None:
        month    = self.month_var.get()
        year     = int(self.year_var.get())
        wdays    = int(self.working_days_var.get())
        uploaded = sum(1 for v in (self.finger_scan_var, self.report_var, self.ptvn_report_var)
                       if v.get())

        self._header_period.configure(text=f"{month}  {year}")
        self._sum_period.set(f"{month} {year}")
        self._sum_wdays.set(f"{wdays} days")
        self._sum_files.set(f"{uploaded} / 3 files")
        self._upload_counter.set(f"{uploaded}/3 uploaded")

        for key, var in (("finger", self.finger_scan_var),
                          ("draft",  self.report_var),
                          ("ptvn",   self.ptvn_report_var)):
            name = Path(var.get()).name if var.get() else None
            self._file_cards[key].set_file(name)

        sub0 = f"{uploaded} file{'s' if uploaded != 1 else ''} ready" if uploaded else "No files uploaded"
        self._check_items[0].update("Upload source files", sub0, done=uploaded > 0)
        self._check_items[1].update("Configure report period", f"{month} {year}", done=True)
        self._check_items[2].update("Set working days", f"{wdays} days selected", done=True)
        gen_sub = ("Completed" if self._is_generated
                   else "In progress…" if self._is_generating else "Pending")
        self._check_items[3].update("Generate attendance report", gen_sub,
                                     done=self._is_generated, active=self._is_generating)

        if self._is_generated:
            self._gen_subtitle.configure(text=f"Completed at {self._gen_time} · Ready to download")
        elif self._is_generating:
            self._gen_subtitle.configure(text="Processing attendance records…")
        elif uploaded == 0:
            self._gen_subtitle.configure(text="Upload at least one file to continue")
        else:
            self._gen_subtitle.configure(
                text=f"{uploaded} file{'s' if uploaded != 1 else ''} ready · Click to generate")

    def _set_progress(self, pct: float) -> None:
        self._progress = max(0.0, min(100.0, pct))
        self._pct_lbl.configure(text=f"{int(self._progress)}%")
        self._prog_fill.place(relwidth=self._progress / 100)
        self._prog_fill.configure(bg=EMERALD if self._is_generated else PRIMARY)

    def _update_badge(self) -> None:
        if self._is_generated:
            self._status_badge.configure(text="● Report Ready", bg=EMERALD_BG, fg=EMERALD_DARK)
        elif self._is_generating:
            self._status_badge.configure(text="● Processing",   bg=PRIMARY_LIGHT, fg=PRIMARY)
        else:
            self._status_badge.configure(text="● Idle",         bg=MUTED_BG,     fg=MUTED_FG)

    def _show_output_row(self, path: Path, ptvn: Path) -> None:
        for w in self._output_frame.winfo_children():
            w.destroy()
        row = tk.Frame(self._output_frame, bg=EMERALD_BG,
                       highlightthickness=1, highlightbackground=EMERALD)
        row.pack(fill="x")
        inner = tk.Frame(row, bg=EMERALD_BG)
        inner.pack(fill="x", padx=12, pady=8)
        ib = tk.Frame(inner, bg="#D1FAE5", width=32, height=32)
        ib.pack_propagate(False)
        ib.pack(side="left", padx=(0, 10))
        tk.Label(ib, text="📄", bg="#D1FAE5", font=("Arial", 14)).pack(expand=True)
        txt = tk.Frame(inner, bg=EMERALD_BG)
        txt.pack(side="left", fill="x", expand=True)
        tk.Label(txt, text=path.name, bg=EMERALD_BG, fg=EMERALD_DARK,
                 font=("Courier", 9, "bold"), anchor="w").pack(fill="x")
        tk.Label(txt, text=ptvn.name, bg=EMERALD_BG, fg=EMERALD,
                 font=("Courier", 9), anchor="w").pack(fill="x")
        ob = tk.Canvas(inner, width=90, height=30, bg=EMERALD_BG,
                       highlightthickness=0, cursor="hand2")
        ob.pack(side="right", padx=(10, 0))
        draw_rounded_rect(ob, 0, 0, 89, 29, 8, fill="#D1FAE5", outline=EMERALD)
        ob.create_text(45, 15, text="↓ Open", fill=EMERALD_DARK, font=("Arial", 9, "bold"))
        ob.bind("<Button-1>", lambda _: self._open_file())

    # ── Traces ────────────────────────────────────────────────────

    def _bind_traces(self) -> None:
        self.month_var.trace_add("write", lambda *_: self._on_period_changed())
        self.year_var.trace_add("write",  lambda *_: self._on_period_changed())
        self.working_days_var.trace_add("write", lambda *_: self._refresh())

    def _on_period_changed(self) -> None:
        try:
            self.working_days_var.set(
                count_weekdays(int(self.year_var.get()), self._selected_month()))
        except Exception:
            pass
        self._refresh()

    def _selected_month(self) -> int:
        return MONTH_NUMBER_BY_NAME.get(self.month_var.get(), date.today().month)

    # ── File pickers ──────────────────────────────────────────────

    def _pick_finger(self) -> None:
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if p:
            self.finger_scan_var.set(p)
            self._refresh()

    def _pick_report(self) -> None:
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if p:
            self.report_var.set(p)
            self._refresh()

    def _pick_ptvn(self) -> None:
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if p:
            self.ptvn_report_var.set(p)
            self._refresh()

    # ── Generate ──────────────────────────────────────────────────

    def _start_processing(self) -> None:
        if self._is_generating:
            return
        uploaded = sum(1 for v in (self.finger_scan_var, self.report_var, self.ptvn_report_var)
                       if v.get())
        if uploaded == 0:
            messagebox.showwarning("No files", "Please upload at least one file.")
            return
        missing = [lbl for lbl, var in (
            ("Report Finger Scan",       self.finger_scan_var),
            ("Draft Attendance",         self.report_var),
            ("Attendance Report (PTVN)", self.ptvn_report_var),
        ) if not var.get()]
        if missing:
            messagebox.showwarning("Missing input", "Please select: " + ", ".join(missing))
            return

        self._is_generating = True
        self._is_generated  = False
        self._set_progress(5)
        self._status_lbl.configure(text="Starting…")
        self._update_badge()
        self._draw_gen_btn(disabled=True)
        self._refresh()
        self.processing_thread = threading.Thread(target=self._process_worker, daemon=True)
        self.processing_thread.start()

    def _process_worker(self) -> None:
        try:
            options = ProcessingOptions(
                month=self._selected_month(),
                year=int(self.year_var.get()),
                fuzzy_threshold=80,
                working_days=int(self.working_days_var.get()),
            )
            output      = self._default_output()
            ptvn_output = self._default_ptvn_output()
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

    def _default_output(self) -> Path:
        base = Path(self.report_var.get()) if self.report_var.get() else Path.home()
        return base.with_name(f"Attendance_Report_{self.month_var.get()}{self.year_var.get()}.xlsx")

    def _default_ptvn_output(self) -> Path:
        ptvn = Path(self.ptvn_report_var.get()) if self.ptvn_report_var.get() else Path.home()
        ext  = ".xlsm" if ptvn.suffix.lower() == ".xlsm" else ".xlsx"
        return ptvn.with_name(f"PTVN_Report_{self.month_var.get()}{self.year_var.get()}{ext}")

    def _drain_queue(self) -> None:
        try:
            while True:
                msg = self.log_queue.get_nowait()
                if msg.startswith("DONE::"):
                    parts = msg.replace("DONE::", "", 1).split("::", 1)
                    out, ptvn = Path(parts[0]), Path(parts[1])
                    self.last_output_path      = out
                    self.last_ptvn_output_path = ptvn
                    self._is_generating = False
                    self._is_generated  = True
                    self._gen_time = datetime.now().strftime("%H:%M")
                    self._set_progress(100)
                    self._status_lbl.configure(text="Done")
                    self._update_badge()
                    self._draw_gen_btn()
                    self._show_output_row(out, ptvn)
                    self._refresh()
                elif msg.startswith("ERROR::"):
                    err = msg.replace("ERROR::", "", 1)
                    self._is_generating = False
                    self._set_progress(0)
                    self._status_lbl.configure(text="Error")
                    self._update_badge()
                    self._draw_gen_btn()
                    self._refresh()
                    messagebox.showerror("Error", err)
                else:
                    pct = min(max(self._progress + 13, 13), 92)
                    self._set_progress(pct)
                    self._status_lbl.configure(text=msg[:60])
                    self._refresh()
        except queue.Empty:
            pass
        self.after(150, self._drain_queue)

    # ── File management ───────────────────────────────────────────

    def _open_file(self) -> None:
        paths = [p for p in (self.last_output_path, self.last_ptvn_output_path)
                 if p and p.exists()]
        if paths:
            for p in paths:
                subprocess.run(["open", str(p)], check=False)
        else:
            messagebox.showinfo("No files", "No generated files are available yet.")

    def _open_folder(self) -> None:
        target = next((p.parent for p in (self.last_output_path, self.last_ptvn_output_path)
                       if p and p.exists()), None)
        if target:
            subprocess.run(["open", str(target)], check=False)
        elif self.finger_scan_var.get():
            subprocess.run(["open", str(Path(self.finger_scan_var.get()).parent)], check=False)
        else:
            messagebox.showinfo("No folder", "No file path available yet.")


def main() -> None:
    app = AttendanceReportApp()
    app.mainloop()


if __name__ == "__main__":
    main()

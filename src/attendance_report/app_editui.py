from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
import queue
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable

from attendance_report.processor import (
    ProcessingOptions,
    process_attendance_report,
)

# ── Palette ────────────────────────────────────────────────────────
BG           = "#F2F5FB"
CARD         = "#FFFFFF"
BORDER       = "#E2E8F0"
BORDER_DASH  = "#C7D2E8"
PRIMARY      = "#3B5BDB"
PRIMARY_FG   = "#FFFFFF"
PRIMARY_LITE = "#EEF2FF"
TEXT         = "#1A1D2E"
MUTED        = "#6B7280"
MUTED_BG     = "#F1F5F9"
EMERALD      = "#059669"
EMERALD_BG   = "#ECFDF5"
EMERALD_DARK = "#047857"
EMERALD_BOX  = "#D1FAE5"
AMBER        = "#D97706"
AMBER_BG     = "#FFFBEB"
GRAY_ICON    = "#94A3B8"
ICON_BOX     = "#EEF2F7"
ROW_BG       = "#F7F9FC"
ROW_BORDER   = "#EDF1F7"
PRIMARY_HOVER= "#2F4CC4"
GEN_OFF_BG   = "#E8ECFB"
GEN_OFF_FG   = "#A9B4E8"
TRACK        = "#E8EDF5"

MONTHS = [
    ("January",1),("February",2),("March",3),("April",4),
    ("May",5),("June",6),("July",7),("August",8),
    ("September",9),("October",10),("November",11),("December",12),
]
MONTH_BY_NUM  = {n: nm for nm, n in MONTHS}
MONTH_BY_NAME = {nm: n  for nm, n in MONTHS}

FILE_CONFIGS = [
    ("finger","Report Finger Scan","Biometric scan export","DRAFT"),
    ("draft", "Draft Attendance",  "Pre-processed draft file","DRAFT"),
    ("ptvn",  "Attendance Report", "Final attendance sheet","FINAL"),
]


def count_weekdays(year: int, month: int) -> int:
    cur = date(year, month, 1)
    end = date(year+1,1,1) if month==12 else date(year,month+1,1)
    n = 0
    while cur < end:
        if cur.weekday() < 5:
            n += 1
        cur += timedelta(days=1)
    return n


FONT    = "Helvetica Neue"   # macOS native; falls back elsewhere
SHADOW  = "#C0CCE0"          # card drop-shadow — dark layer
SHADOW2 = "#D4DFEF"          # card drop-shadow — light layer


def _lerp_color(a: str, b: str, t: float) -> str:
    """Linearly interpolate between two hex colours (0.0=a, 1.0=b)."""
    ra, ga, ba_ = int(a[1:3],16), int(a[3:5],16), int(a[5:],16)
    rb, gb, bb  = int(b[1:3],16), int(b[3:5],16), int(b[5:],16)
    return "#{:02x}{:02x}{:02x}".format(
        int(ra+(rb-ra)*t), int(ga+(gb-ga)*t), int(ba_+(bb-ba_)*t))


def rrect(cv: tk.Canvas, x1: int, y1: int, x2: int, y2: int,
          r: int, **kw: Any) -> int:
    pts = [x1+r,y1, x2-r,y1, x2,y1, x2,y1+r,
           x2,y2-r, x2,y2, x2-r,y2, x1+r,y2,
           x1,y2, x1,y2-r, x1,y1+r, x1,y1]
    return cv.create_polygon(pts, smooth=True, splinesteps=36, **kw)


def tag_rect(cv: tk.Canvas, x1: int, y1: int, x2: int, y2: int,
             r: int, **kw: Any) -> int:
    """Rounded rect with top-right corner left sharp (tag / chip shape)."""
    pts = [
        x1+r, y1,                   # top-left arc start
        x2,   y1,                   # top-right SHARP → no arc
        x2,   y1,
        x2,   y2-r,                 # right side
        x2,   y2,  x2-r, y2,        # bottom-right arc
        x1+r, y2,  x1,   y2,        # bottom-left arc
        x1,   y1+r, x1,  y1,        # left side back up
    ]
    return cv.create_polygon(pts, smooth=True, splinesteps=36, **kw)


def draw_icon(cv: tk.Canvas, kind: str, x: int, y: int, s: int, fg: str) -> None:
    if kind == "finger":
        # Scan bracket corners + 3 fingerprint arcs
        m  = int(s * 0.37)
        bk = max(5, m // 2)
        for sx, sy in ((-1,-1), (1,-1), (1,1), (-1,1)):
            bx, by = x + sx*m, y + sy*m
            cv.create_line(bx, by, bx, by - sy*bk,
                           fill=fg, width=2.0, capstyle="round")
            cv.create_line(bx, by, bx - sx*bk, by,
                           fill=fg, width=2.0, capstyle="round")
        for r in (4, 7, 11):
            cv.create_arc(x-r, y-r+1, x+r, y+r+1,
                          start=20, extent=140,
                          style="arc", outline=fg, width=1.5)
    elif kind == "draft":
        # Document with dog-eared top-right corner + 3 body lines
        pw   = int(s * 0.28)
        fold = max(5, pw // 2)
        ph   = int(s * 0.44)
        pts  = [x-pw, y-ph,  x+pw-fold, y-ph,
                x+pw, y-ph+fold,  x+pw, y+ph,  x-pw, y+ph]
        cv.create_polygon(pts, outline=fg, fill="", width=1.6)
        cv.create_line(x+pw-fold, y-ph, x+pw-fold, y-ph+fold,
                       x+pw, y-ph+fold, fill=fg, width=1.2)
        lx = int(pw * 0.72)
        for dy in (-ph//3, ph//8, ph//2):
            cv.create_line(x-lx, y+dy, x+lx, y+dy,
                           fill=fg, width=1.3, capstyle="round")
    elif kind == "ptvn":
        # Spreadsheet / grid icon (2x3 cells)
        pw = int(s * 0.30)
        cv.create_rectangle(x-pw, y-pw, x+pw, y+pw, outline=fg, width=1.6)
        cv.create_line(x, y-pw, x, y+pw, fill=fg, width=1.1)
        d = pw // 3
        cv.create_line(x-pw, y-d, x+pw, y-d, fill=fg, width=1.1)
        cv.create_line(x-pw, y+d, x+pw, y+d, fill=fg, width=1.1)


def draw_metric_icon(cv: tk.Canvas, kind: str,
                     cx: int, cy: int, r: int, fg: str) -> None:
    """Geometric metric icons — no emoji, clean canvas-drawn shapes."""
    if kind == "period":
        # Calendar: rect + header bar + day dots
        cv.create_rectangle(cx-r, cy-r+1, cx+r, cy+r,
                            outline=fg, width=1.5)
        cv.create_line(cx-r, cy-r+6, cx+r, cy-r+6, fill=fg, width=1.3)
        for hx in (cx-4, cx+4):
            cv.create_line(hx, cy-r-1, hx, cy-r+4,
                           fill=fg, width=2.0, capstyle="round")
        for dx in (-4, 0, 4):
            for dy_off in (3, 8):
                yo = cy - r + 6 + dy_off
                cv.create_oval(cx+dx-1.5, yo-1.5,
                               cx+dx+1.5, yo+1.5, fill=fg, outline="")
    elif kind == "days":
        # Clock face: circle + hour + minute hands + centre dot
        cv.create_oval(cx-r+1, cy-r+1, cx+r-1, cy+r-1,
                       outline=fg, width=1.5)
        cv.create_line(cx, cy, cx, cy-r+4,
                       fill=fg, width=2, capstyle="round")
        cv.create_line(cx, cy, cx+r-5, cy,
                       fill=fg, width=2, capstyle="round")
        cv.create_oval(cx-2, cy-2, cx+2, cy+2, fill=fg, outline="")
    elif kind == "files":
        # Bar chart: 3 columns of varying height
        bw, gap = 4, 3
        fracs   = [0.50, 0.92, 0.68]
        max_h   = r * 2 - 2
        x0      = cx - (bw * 3 + gap * 2) // 2
        base    = cy + r
        for i, frac in enumerate(fracs):
            bx = x0 + i * (bw + gap)
            bh = int(max_h * frac)
            cv.create_rectangle(bx, base - bh, bx + bw, base,
                               fill=fg, outline="")
        cv.create_line(cx - r, base + 1, cx + r, base + 1,
                       fill=fg, width=1.3)


# ── Card ───────────────────────────────────────────────────────────
class Card(tk.Frame):
    """White rounded card. A background canvas paints the rounded rect and
    the content `body` is packed on top so it always renders."""
    def __init__(self, parent: tk.Misc, *,
                 pad: tuple[int,int,int,int] = (28,24,28,24),
                 r: int = 22, bg: str = BG,
                 fill: str = CARD, outline: str = BORDER,
                 ow: float = 0.8) -> None:
        super().__init__(parent, bg=bg)
        self._r, self._fill, self._outline, self._ow = r, fill, outline, ow
        self._bgc = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self._bgc.place(x=0, y=0, relwidth=1, relheight=1)
        pl, pt, pr, pb = pad
        self.body = tk.Frame(self, bg=fill)
        self.body.pack(fill="both", expand=True, padx=(pl, pr), pady=(pt, pb))
        self.bind("<Configure>", self._redraw)

    def _redraw(self, _evt: object = None) -> None:
        w, h = self.winfo_width(), self.winfo_height()
        if w < 4 or h < 4:
            return
        self._bgc.delete("all")
        # two-layer drop-shadow (offset right+2, down+4)
        rrect(self._bgc, 3, 5, w-1, h-1, self._r, fill=SHADOW,  outline="")
        rrect(self._bgc, 2, 3, w-2, h-2, self._r, fill=SHADOW2, outline="")
        # card face
        rrect(self._bgc, 1, 1, w-4, h-5, self._r,
              fill=self._fill, outline=self._outline, width=self._ow)


# ── File upload card ───────────────────────────────────────────────
class FileCard(tk.Frame):
    """Dashed rounded upload tile. The background canvas paints the dashed
    rounded border; the content `body` is packed on top."""
    def __init__(self, parent: tk.Misc, key: str, label: str,
                 desc: str, badge: str,
                 command: Callable[[], None]) -> None:
        super().__init__(parent, bg=BG, cursor="hand2")
        self._key   = key
        self._label = label
        self._desc  = desc
        self._badge = badge
        self._cmd   = command
        self._file: str | None = None
        self._bgc = tk.Canvas(self, bg=BG, highlightthickness=0, bd=0,
                              cursor="hand2")
        self._bgc.place(x=0, y=0, relwidth=1, relheight=1)
        self.body = tk.Frame(self, bg=CARD, cursor="hand2")
        self.body.pack(fill="both", expand=True, padx=20, pady=20)
        self._hover_t   = 0.0
        self._hover_aid: str | None = None
        self.bind("<Configure>", lambda _e: self._redraw())
        self._bgc.bind("<Button-1>", lambda _: command())
        self._rebuild()
        self.bind("<Enter>", lambda _: self._fc_animate(1))
        self.bind("<Leave>", lambda _: self._fc_animate(-1))

    def _redraw(self) -> None:
        w, h = self.winfo_width(), self.winfo_height()
        if w < 4 or h < 4:
            return
        self._bgc.delete("all")
        ready = self._file is not None
        base  = EMERALD_BG if ready else CARD
        tint  = EMERALD_BOX if ready else "#EDF1FF"
        fill  = _lerp_color(base, tint,
                            getattr(self, "_hover_t", 0.0) * 0.5)
        rrect(self._bgc, 2, 2, w-2, h-2, 24,
              fill=fill,
              outline=EMERALD if ready else BORDER_DASH,
              width=1.6, dash=None if ready else (5, 4))

    def _bind_click(self, widget: tk.Misc) -> None:
        widget.bind("<Button-1>", lambda _: self._cmd())
        for ch in widget.winfo_children():
            self._bind_click(ch)

    def _rebuild(self) -> None:
        for w in self.body.winfo_children():
            w.destroy()
        ready = self._file is not None
        bg = EMERALD_BG if ready else CARD
        self.body.configure(bg=bg)

        top = tk.Frame(self.body, bg=bg)
        top.pack(fill="x")
        box_bg = EMERALD_BOX if ready else ICON_BOX
        icv = tk.Canvas(top, width=52, height=52, bg=bg, highlightthickness=0)
        icv.pack(side="left")
        icv.create_oval(2, 2, 50, 50, fill=box_bg, outline="")
        draw_icon(icv, self._key, 26, 26, 34, EMERALD if ready else MUTED)
        badge_bg = PRIMARY_LITE if self._badge == "FINAL" else MUTED_BG
        badge_fg = PRIMARY      if self._badge == "FINAL" else MUTED
        tk.Label(top, text=self._badge, bg=badge_bg, fg=badge_fg,
                 font=(FONT,8,"bold"), padx=8, pady=4
                 ).pack(side="right", anchor="n")

        tk.Label(self.body, text=self._label, bg=bg, fg=TEXT,
                 font=(FONT,11,"bold"), anchor="w").pack(fill="x",
                 pady=(16,0))
        tk.Label(self.body, text=self._desc, bg=bg, fg=MUTED,
                 font=(FONT,9), anchor="w").pack(fill="x", pady=(2,14))
        if ready:
            tk.Label(self.body, text=f"Selected \u2014 {self._file}", bg=bg,
                     fg=EMERALD_DARK, font=(FONT,9,"bold"),
                     anchor="w", wraplength=180,
                     justify="left").pack(fill="x")
        else:
            tk.Label(self.body, text="Select file  \u2192",
                     bg=bg, fg=MUTED, font=(FONT,9),
                     anchor="w").pack(fill="x")
        self._bind_click(self.body)

    def set_file(self, name: str | None) -> None:
        self._file = name
        self._rebuild()
        self._redraw()

    def _fc_animate(self, direction: int) -> None:
        if self._hover_aid:
            self.after_cancel(self._hover_aid)
            self._hover_aid = None
        def _tick() -> None:
            self._hover_t = max(0.0, min(1.0,
                                         self._hover_t + direction * 0.25))
            self._redraw()
            if 0.0 < self._hover_t < 1.0:
                self._hover_aid = self.after(14, _tick)
        _tick()


# ── Summary row ────────────────────────────────────────────────────
class SummaryRow(tk.Frame):
    def __init__(self, parent: tk.Misc, icon: str, label: str,
                 var: tk.StringVar,
                 icon_bg: str = PRIMARY_LITE,
                 icon_fg: str = PRIMARY) -> None:
        super().__init__(parent, bg=CARD)
        self._bgc = tk.Canvas(self, bg=CARD, highlightthickness=0, bd=0)
        self._bgc.place(x=0, y=0, relwidth=1, relheight=1)
        inner = tk.Frame(self, bg=ROW_BG)
        inner.pack(fill="both", expand=True, padx=14, pady=13)
        ic = tk.Canvas(inner, width=42, height=42, bg=ROW_BG,
                       highlightthickness=0)
        ic.pack(side="left", padx=(0,14))
        ic.create_oval(2, 2, 40, 40, fill=icon_bg, outline="")
        draw_metric_icon(ic, icon, 21, 21, 9, icon_fg)
        txt = tk.Frame(inner, bg=ROW_BG)
        txt.pack(side="left", fill="x", expand=True)
        tk.Label(txt, text=label.upper(), bg=ROW_BG, fg=MUTED,
                 font=(FONT, 8, "bold")).pack(anchor="w")
        tk.Label(txt, textvariable=var, bg=ROW_BG, fg=TEXT,
                 font=(FONT, 14, "bold")).pack(anchor="w")
        self.bind("<Configure>", self._redraw)

    def _redraw(self, _evt: object = None) -> None:
        w, h = self.winfo_width(), self.winfo_height()
        if w < 4 or h < 4:
            return
        self._bgc.delete("all")
        rrect(self._bgc, 1, 1, w-1, h-1, 13,
              fill=ROW_BG, outline=ROW_BORDER, width=1)


# ── Checklist item ─────────────────────────────────────────────────
class CheckItem(tk.Frame):
    def __init__(self, parent: tk.Misc, label: str, sub: str,
                 done: bool = False) -> None:
        super().__init__(parent, bg=CARD)
        self._cv = tk.Canvas(self, width=22, height=22, bg=CARD,
                             highlightthickness=0)
        self._cv.pack(side="left", anchor="n", padx=(0,12), pady=1)
        right = tk.Frame(self, bg=CARD)
        right.pack(side="left", fill="x", expand=True)
        self._lbl = tk.Label(right, text=label, bg=CARD,
                             font=(FONT,10), fg=TEXT)
        self._lbl.pack(anchor="w")
        self._sub = tk.Label(right, text=sub, bg=CARD, fg=MUTED,
                             font=(FONT,9))
        self._sub.pack(anchor="w")
        self._paint(done, False)

    def _paint(self, done: bool, active: bool) -> None:
        self._cv.delete("all")
        if done:
            self._cv.create_oval(2,2,20,20, fill=EMERALD, outline=EMERALD)
            self._cv.create_text(11,11, text="\u2713", fill=PRIMARY_FG,
                                 font=(FONT,9,"bold"))
        elif active:
            self._cv.create_oval(2,2,20,20, fill=PRIMARY_LITE,
                                 outline=PRIMARY, width=2)
            self._cv.create_oval(7,7,15,15, fill=PRIMARY, outline="")
        else:
            self._cv.create_oval(2,2,20,20, fill=CARD,
                                 outline="#CBD5E1", width=2)

    def update(self, label: str, sub: str,
               done: bool, active: bool = False) -> None:
        self._paint(done, active)
        self._lbl.configure(
            text=label,
            fg=TEXT if done else (PRIMARY if active else MUTED),
            font=(FONT,10,"bold" if done else "normal"))
        self._sub.configure(text=sub)


# ── Rounded field (combobox container) ────────────────────────────
class RoundedField(tk.Frame):
    """Canvas-drawn rounded border that wraps a flat ttk.Combobox."""
    def __init__(self, parent: tk.Misc, r: int = 14) -> None:
        super().__init__(parent, bg=CARD)
        self._r       = r
        self._focused = False
        self._bgc = tk.Canvas(self, bg=CARD, highlightthickness=0, bd=0)
        self._bgc.place(x=0, y=0, relwidth=1, relheight=1)
        self.bind("<Configure>", self._redraw)

    def focus_in(self) -> None:
        self._focused = True;  self._redraw()

    def focus_out(self) -> None:
        self._focused = False; self._redraw()

    def _redraw(self, _evt: object = None) -> None:
        w, h = self.winfo_width(), self.winfo_height()
        if w < 4 or h < 4:
            return
        self._bgc.delete("all")
        col, lw = (PRIMARY, 1.8) if self._focused else (BORDER, 1.3)
        rrect(self._bgc, 1, 1, w-1, h-1, self._r,
              fill=CARD, outline=col, width=lw)


# ── Modern dropdown ────────────────────────────────────────────────
class ModernDropdown(tk.Frame):
    """Canvas trigger + floating Toplevel popup with custom scrollbar."""

    ITEM_H   = 42
    POPUP_R  = 16
    FIELD_R  = 12
    SEL_BG   = "#EEF4FF"
    HOV_BG   = "#E7F0FF"
    SCR_W    = 6
    TRK_C    = "#F1F5F9"
    THB_C    = "#C7D7FE"
    MAX_ROWS = 8

    def __init__(self, parent: tk.Misc,
                 textvariable: tk.Variable,
                 values: list,
                 font: tuple = (FONT, 11)) -> None:
        super().__init__(parent, bg=CARD)
        self._var   = textvariable
        self._vals  = [str(v) for v in values]
        self._fnt   = font
        self._open  = False
        self._pop:  tk.Toplevel | None = None
        self._hidx  = -1
        self._sfrac = 0.0
        self._ct    = 0.0
        self._caid: str | None = None
        self._dyref: float | None = None
        self._obid:  str | None = None
        self._cv = tk.Canvas(self, height=48, bg=CARD,
                             highlightthickness=0, cursor="hand2")
        self._cv.pack(fill="x")
        self._cv.bind("<Configure>", lambda _: self._drw_trig())
        self._cv.bind("<Button-1>",  lambda _: self._toggle())
        self._var.trace_add("write", lambda *_: self._drw_trig())

    # ── trigger ───────────────────────────────────────────────────
    def _drw_trig(self) -> None:
        cv = self._cv
        w  = cv.winfo_width()
        if w < 4:
            return
        cv.delete("all")
        col = PRIMARY if self._open else BORDER
        lw  = 1.8    if self._open else 1.3
        rrect(cv, 1, 1, w-1, 47, self.FIELD_R,
              fill=CARD, outline=col, width=lw)
        cv.create_text(18, 24, text=str(self._var.get()),
                       anchor="w", fill=TEXT, font=self._fnt)
        cx, cy = w-22, 24
        t = self._ct
        hv, arm = 3, 6
        tip_y  = cy + hv * (1 - 2*t)
        base_y = cy - hv * (1 - 2*t)
        cv.create_line(cx-arm, base_y, cx, tip_y,
                       fill=MUTED, width=1.8,
                       capstyle="round", joinstyle="round")
        cv.create_line(cx, tip_y, cx+arm, base_y,
                       fill=MUTED, width=1.8,
                       capstyle="round", joinstyle="round")

    def _anim(self, tgt: float) -> None:
        if self._caid:
            self.after_cancel(self._caid)
            self._caid = None
        def _t() -> None:
            d = tgt - self._ct
            if abs(d) < 0.04:
                self._ct = tgt; self._drw_trig(); return
            self._ct += d * 0.35; self._drw_trig()
            self._caid = self.after(12, _t)
        _t()

    def _toggle(self) -> None:
        self._close() if self._open else self._show()

    # ── popup ─────────────────────────────────────────────────────
    def _show(self) -> None:
        if self._pop:
            return
        self.update_idletasks()
        rx = self._cv.winfo_rootx()
        ry = self._cv.winfo_rooty() + self._cv.winfo_height() + 4
        pw = self._cv.winfo_width()
        n   = len(self._vals)
        vis = min(n, self.MAX_ROWS)
        pv  = 8
        ph  = vis * self.ITEM_H + pv * 2
        cur = str(self._var.get())
        try:
            si = self._vals.index(cur)
        except ValueError:
            si = 0
        ms = max(0, n - vis)
        self._sfrac = max(0.0, min(1.0, (si - vis//2) / ms)) if ms else 0.0
        pop = tk.Toplevel(self)
        pop.wm_overrideredirect(True)
        pop.wm_geometry(f"{pw}x{ph}+{rx}+{ry}")
        pop.configure(bg=CARD)
        pop.attributes("-topmost", True)
        self._pop = pop
        self._vis = vis
        self._ph  = ph
        self._pw  = pw
        bgc = tk.Canvas(pop, bg=CARD, highlightthickness=0)
        bgc.place(x=0, y=0, relwidth=1, relheight=1)
        bgc.bind("<Configure>", lambda e, c=bgc:
                 (c.delete("all"),
                  rrect(c, 1, 1, e.width-2, e.height-2,
                        self.POPUP_R, fill=CARD,
                        outline=BORDER, width=1.0)))
        icw = pw - self.SCR_W - 7
        ich = vis * self.ITEM_H
        ic  = tk.Canvas(pop, bg=CARD, highlightthickness=0)
        ic.place(x=3, y=pv, width=icw, height=ich)
        self._ic  = ic
        self._icw = icw
        self._ich = ich
        sc = tk.Canvas(pop, bg=self.TRK_C, highlightthickness=0)
        sc.place(x=pw-self.SCR_W-2, y=pv,
                 width=self.SCR_W, height=ich)
        self._sc  = sc
        self._sch = ich
        self._drw_items()
        self._drw_scroll()
        ic.bind("<Motion>",          self._on_hov)
        ic.bind("<Leave>",           lambda _: self._set_hov(-1))
        ic.bind("<Button-1>",        self._on_clk)
        ic.bind("<MouseWheel>",      self._on_whl)
        pop.bind("<MouseWheel>",     self._on_whl)
        sc.bind("<Button-1>",        self._on_sc_dn)
        sc.bind("<B1-Motion>",       self._on_sc_mv)
        sc.bind("<ButtonRelease-1>", lambda _: setattr(self, "_dyref", None))
        self._open = True
        self._anim(1.0)
        self._obid = self.winfo_toplevel().bind(
            "<Button-1>", self._on_out, add="+")

    def _soff(self) -> int:
        return int(self._sfrac * max(0, len(self._vals) - self._vis)
                   * self.ITEM_H)

    def _drw_items(self) -> None:
        cv  = self._ic
        cv.delete("all")
        ih  = self.ITEM_H
        off = self._soff()
        cur = str(self._var.get())
        for i, v in enumerate(self._vals):
            y0 = i*ih - off
            y1 = y0 + ih
            if y1 <= 0 or y0 >= self._ich:
                continue
            if v == cur:
                bg = self.SEL_BG
            elif i == self._hidx:
                bg = self.HOV_BG
            else:
                bg = CARD
            if bg != CARD:
                rrect(cv, 5, y0+3, self._icw-5, y1-3, 8,
                      fill=bg, outline="")
            cv.create_text(
                20, y0+ih//2, text=v, anchor="w", fill=TEXT,
                font=(self._fnt[0], self._fnt[1],
                      "bold" if v == cur else "normal"))

    def _drw_scroll(self) -> None:
        cv = self._sc
        cv.delete("all")
        n, vis = len(self._vals), self._vis
        if n <= vis:
            return
        h  = self._sch
        th = max(24, int(h * vis / n))
        ty = int(self._sfrac * (h - th))
        r  = self.SCR_W // 2
        rrect(cv, 0, ty, self.SCR_W, ty+th, r,
              fill=self.THB_C, outline="")

    def _set_hov(self, idx: int) -> None:
        if idx != self._hidx:
            self._hidx = idx
            self._drw_items()

    def _on_hov(self, e: tk.Event) -> None:
        off = self._soff()
        idx = (e.y + off) // self.ITEM_H
        self._set_hov(idx if 0 <= idx < len(self._vals) else -1)

    def _on_clk(self, e: tk.Event) -> None:
        off = self._soff()
        idx = (e.y + off) // self.ITEM_H
        if 0 <= idx < len(self._vals):
            v = self._vals[idx]
            try:
                self._var.set(int(v))
            except (ValueError, TypeError):
                self._var.set(v)
            self._close()

    def _on_whl(self, e: tk.Event) -> None:
        n, vis = len(self._vals), self._vis
        if n <= vis:
            return
        step = 3 / (n - vis)
        self._sfrac = max(0.0, min(1.0,
            self._sfrac + (step if e.delta < 0 else -step)))
        self._drw_items()
        self._drw_scroll()

    def _on_sc_dn(self, e: tk.Event) -> None:
        self._dyref = e.y

    def _on_sc_mv(self, e: tk.Event) -> None:
        if self._dyref is None:
            return
        n, vis = len(self._vals), self._vis
        if n <= vis:
            return
        th = max(24, int(self._sch * vis / n))
        mx = self._sch - th
        if mx <= 0:
            return
        self._sfrac = max(0.0, min(1.0,
            self._sfrac + (e.y - self._dyref) / mx))
        self._dyref = e.y
        self._drw_items()
        self._drw_scroll()

    def _on_out(self, e: tk.Event) -> None:
        if not self._pop:
            return
        try:
            tx = self._cv.winfo_rootx()
            ty = self._cv.winfo_rooty()
            if (tx-2 <= e.x_root <= tx + self._cv.winfo_width() + 2 and
                    ty-2 <= e.y_root <= ty + self._cv.winfo_height() + 2):
                return
        except Exception:
            pass
        self._close()

    def _close(self) -> None:
        if self._obid:
            try:
                self.winfo_toplevel().unbind("<Button-1>", self._obid)
            except Exception:
                pass
            self._obid = None
        if self._pop:
            try:
                self._pop.destroy()
            except Exception:
                pass
            self._pop = None
        self._hidx = -1
        self._open = False
        self._anim(0.0)


# ── Outline button ─────────────────────────────────────────────────
class OutlineBtn(tk.Frame):
    def __init__(self, parent: tk.Misc, text: str,
                 cmd: Callable[[], None],
                 w: int = 190, h: int = 38) -> None:
        super().__init__(parent, bg=CARD)
        # NB: never use self._w / self._h — tkinter uses _w as the widget path.
        self._bw, self._bh, self._txt = w, h, text
        self._ht:  float      = 0.0
        self._aid: str | None = None
        self._cv = tk.Canvas(self, width=w, height=h, bg=CARD,
                             highlightthickness=0, cursor="hand2")
        self._cv.pack()
        self._cv.bind("<Button-1>", lambda _: cmd())
        self._cv.bind("<Enter>",    lambda _: self._animate(1))
        self._cv.bind("<Leave>",    lambda _: self._animate(-1))
        self._paint()

    def _animate(self, direction: int) -> None:
        if self._aid:
            self.after_cancel(self._aid)
            self._aid = None
        def _tick() -> None:
            self._ht = max(0.0, min(1.0, self._ht + direction * 0.2))
            self._paint()
            if 0.0 < self._ht < 1.0:
                self._aid = self.after(14, _tick)
        _tick()

    def _paint(self) -> None:
        cv = self._cv
        cv.delete("all")
        bg = _lerp_color(CARD, MUTED_BG, self._ht)
        rrect(cv, 1, 1, self._bw-1, self._bh-1, 10,
              fill=bg, outline=BORDER, width=1)
        cv.create_text(self._bw//2, self._bh//2, text=self._txt,
                       fill=TEXT, font=(FONT, 10, "bold"))


# ── App ────────────────────────────────────────────────────────────
class AttendanceReportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Attendance Report")
        self.geometry("1440x900")
        self.minsize(1140, 720)
        self.configure(bg=BG)

        self.log_q:   queue.Queue[str] = queue.Queue()
        self.thread:  threading.Thread | None = None
        self.out_path: Path | None = None
        self.ptvn_out: Path | None = None

        today = date.today()
        self.finger_var = tk.StringVar()
        self.draft_var  = tk.StringVar()
        self.ptvn_var   = tk.StringVar()
        self.month_var  = tk.StringVar(value=MONTH_BY_NUM[today.month])
        self.year_var   = tk.IntVar(value=today.year)
        self.wdays_var  = tk.IntVar(
            value=count_weekdays(today.year, today.month))
        self._sum_period = tk.StringVar()
        self._sum_wdays  = tk.StringVar()
        self._sum_files  = tk.StringVar()

        self._generating  = False
        self._generated   = False
        self._progress    = 0.0
        self._gen_time    = ""
        self._gen_ht      = 0.0
        self._gen_aid: str | None = None
        self._badge_state = ("\u25cf  Idle", MUTED_BG, MUTED)

        self._build()
        self._trace()
        self._refresh()
        self.after(150, self._poll)

    # ── Layout ────────────────────────────────────────────────────
    def _build(self) -> None:
        self._setup_style()
        self._header()
        self._main()
        self._footer()

    def _setup_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("App.TCombobox",
                        fieldbackground=CARD, background=CARD,
                        foreground=TEXT, arrowcolor=MUTED,
                        bordercolor=BORDER, lightcolor=BORDER,
                        darkcolor=BORDER, relief="flat", padding=7)
        style.map("App.TCombobox",
                  fieldbackground=[("readonly", CARD)],
                  selectbackground=[("readonly", CARD)],
                  selectforeground=[("readonly", TEXT)],
                  bordercolor=[("focus", PRIMARY), ("active", PRIMARY)],
                  arrowcolor=[("active", PRIMARY)])
        self.option_add("*TCombobox*Listbox.background", CARD)
        self.option_add("*TCombobox*Listbox.foreground", TEXT)
        self.option_add("*TCombobox*Listbox.selectBackground", PRIMARY_LITE)
        self.option_add("*TCombobox*Listbox.selectForeground", TEXT)
        self.option_add("*TCombobox*Listbox.font", "Arial 11")
        self.option_add("*TCombobox*Listbox.borderWidth", 0)
        self.option_add("*TCombobox*Listbox.relief", "flat")
        # Modern slim scrollbar — used by the combobox dropdown listbox
        style.configure("Vertical.TScrollbar",
                        gripcount=0, arrowsize=0,
                        background=MUTED_BG, troughcolor=ROW_BG,
                        bordercolor=CARD, lightcolor=CARD,
                        darkcolor=CARD, relief="flat",
                        width=6)
        style.map("Vertical.TScrollbar",
                  background=[("active", BORDER), ("!active", MUTED_BG)])
        style.layout("Vertical.TScrollbar",
                     [("Vertical.TScrollbar.trough",
                       {"sticky": "ns",
                        "children": [("Vertical.TScrollbar.thumb",
                                      {"expand": "1",
                                       "sticky": "nswe"})]})])
        # borderless variant — used inside RoundedField wrappers
        style.configure("Flat.TCombobox",
                        fieldbackground=CARD, background=CARD,
                        foreground=TEXT, arrowcolor=MUTED,
                        bordercolor=CARD, lightcolor=CARD,
                        darkcolor=CARD, relief="flat", padding=(14, 9))
        style.map("Flat.TCombobox",
                  fieldbackground=[("readonly", CARD)],
                  selectbackground=[("readonly", CARD)],
                  selectforeground=[("readonly", TEXT)],
                  arrowcolor=[("active", PRIMARY)])

    def _header(self) -> None:
        hdr = tk.Frame(self, bg=CARD, height=68)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        logo = tk.Canvas(hdr, width=40, height=40, bg=CARD,
                         highlightthickness=0)
        logo.pack(side="left", padx=(32,12), pady=14)
        rrect(logo, 1, 1, 39, 39, 11, fill=PRIMARY, outline="")
        logo.create_polygon(20,8, 30,14, 20,20, 10,14,
                            fill=PRIMARY_FG, outline="")
        logo.create_line(10,20, 20,26, 30,20,
                         fill=PRIMARY_FG, width=2, joinstyle="round")
        logo.create_line(10,25, 20,31, 30,25,
                         fill=PRIMARY_FG, width=2, joinstyle="round")

        tf = tk.Frame(hdr, bg=CARD)
        tf.pack(side="left")
        tk.Label(tf, text="Attendance Report", bg=CARD, fg=TEXT,
                 font=(FONT,14,"bold")).pack(anchor="w")
        tk.Label(tf, text="HR Reporting System", bg=CARD, fg=MUTED,
                 font=(FONT,9)).pack(anchor="w")

        # period label and badge removed — header right side is clean
        self._badge_cv   = tk.Canvas(hdr, width=0, height=0,
                                     highlightthickness=0)   # kept for _draw_badge/_update_badge compat
        self._period_lbl = tk.Label(hdr, bg=CARD)            # kept for _refresh compat

    def _main(self) -> None:
        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True, padx=32, pady=24)
        outer.columnconfigure(0, weight=1)
        outer.columnconfigure(1, minsize=300)
        outer.rowconfigure(0, weight=1)

        left = tk.Frame(outer, bg=BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0,18))
        self._upload_card(left)
        self._period_card(left)
        self._generate_card(left)
        self._file_mgmt_card(left)

        right = tk.Frame(outer, bg=BG)
        right.grid(row=0, column=1, sticky="new")
        self._summary_panel(right)
        self._checklist_panel(right)

    # ── Cards ─────────────────────────────────────────────────────
    def _step_badge_widget(self, parent: tk.Frame, step: str) -> None:
        cv = tk.Canvas(parent, width=30, height=30, bg=CARD,
                       highlightthickness=0)
        cv.pack(side="left", padx=(0,12))
        cv.create_oval(1, 1, 29, 29, fill=PRIMARY, outline="")
        cv.create_text(15, 15, text=step, fill=PRIMARY_FG,
                       font=(FONT,12,"bold"))

    def _card_header(self, parent: tk.Frame, step: str, title: str,
                     right_var: tk.StringVar | None = None) -> None:
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill="x", pady=(0,18))
        self._step_badge_widget(row, step)
        tk.Label(row, text=title, bg=CARD, fg=TEXT,
                 font=(FONT,14,"bold")).pack(side="left")
        if right_var is not None:
            tk.Label(row, textvariable=right_var, bg=CARD, fg=MUTED,
                     font=(FONT,10)).pack(side="right")

    def _upload_card(self, parent: tk.Frame) -> None:
        card = Card(parent)
        card.pack(fill="x", pady=(0,12))
        self._upload_ctr = tk.StringVar(value="0/3 uploaded")
        self._card_header(card.body, "1",
                          "Upload Attendance Files", self._upload_ctr)
        grid = tk.Frame(card.body, bg=CARD)
        grid.pack(fill="x")
        for c in range(3):
            grid.columnconfigure(c, weight=1, uniform="fc")
        pmap = {"finger": self._pick_finger,
                "draft":  self._pick_draft,
                "ptvn":   self._pick_ptvn}
        self._fcards: dict[str, FileCard] = {}
        for col, (key, label, desc, badge) in enumerate(FILE_CONFIGS):
            fc = FileCard(grid, key, label, desc, badge, pmap[key])
            fc.grid(row=0, column=col, sticky="nsew",
                    padx=(0 if col == 0 else 10, 0))
            self._fcards[key] = fc

    def _period_card(self, parent: tk.Frame) -> None:
        card = Card(parent)
        card.pack(fill="x", pady=(0,12))
        self._card_header(card.body, "2", "Select Report Period")
        grid = tk.Frame(card.body, bg=CARD)
        grid.pack(fill="x")
        for c in range(3):
            grid.columnconfigure(c, weight=1, uniform="ps")
        fields = [
            ("MONTH",        self.month_var, [nm for nm,_ in MONTHS]),
            ("YEAR",         self.year_var,  list(range(2020,2102))),
            ("WORKING DAYS", self.wdays_var, list(range(1,32))),
        ]
        for col, (lbl, var, vals) in enumerate(fields):
            cf = tk.Frame(grid, bg=CARD)
            cf.grid(row=0, column=col, sticky="ew",
                    padx=(0 if col == 0 else 12, 0))
            tk.Label(cf, text=lbl, bg=CARD, fg=MUTED,
                     font=(FONT,9,"bold")).pack(anchor="w", pady=(0,6))
            ModernDropdown(cf, textvariable=var,
                           values=vals).pack(fill="x")

    def _generate_card(self, parent: tk.Frame) -> None:
        card = Card(parent)
        card.pack(fill="x", pady=(0,12))
        b = card.body

        top = tk.Frame(b, bg=CARD)
        top.pack(fill="x", pady=(0,14))
        lf = tk.Frame(top, bg=CARD)
        lf.pack(side="left", fill="x", expand=True)
        tk.Label(lf, text="Generate Report", bg=CARD, fg=TEXT,
                 font=(FONT,14,"bold")).pack(anchor="w")
        self._gen_sub = tk.Label(lf, bg=CARD, fg=MUTED, font=(FONT,10),
                                 text="Upload at least one file to continue")
        self._gen_sub.pack(anchor="w", pady=(3,0))

        self._gen_cv = tk.Canvas(top, width=168, height=46, bg=CARD,
                                 highlightthickness=0, cursor="hand2")
        self._gen_cv.pack(side="right", padx=(16,0))
        self._gen_cv.bind("<Button-1>", lambda _: self._start())
        self._gen_cv.bind("<Enter>",    lambda _: self._gen_animate(1))
        self._gen_cv.bind("<Leave>",    lambda _: self._gen_animate(-1))
        self._draw_gen_btn()

        sr = tk.Frame(b, bg=CARD)
        sr.pack(fill="x", pady=(0,8))
        self._status_lbl = tk.Label(sr, text="Awaiting input",
                                    bg=CARD, fg=MUTED,
                                    font=(FONT, 10))
        self._status_lbl.pack(side="left")
        self._pct_lbl = tk.Label(sr, text="0%", bg=CARD, fg=TEXT,
                                 font=(FONT, 10, "bold"))
        self._pct_lbl.pack(side="right")

        self._prog_cv = tk.Canvas(b, height=10, bg=CARD,
                                  highlightthickness=0)
        self._prog_cv.pack(fill="x", pady=(0,2))
        self._prog_cv.bind("<Configure>", lambda _: self._draw_progress())

        self._out_frame = tk.Frame(b, bg=CARD)
        self._out_frame.pack(fill="x", pady=(10,0))

    def _gen_animate(self, direction: int) -> None:
        if self._gen_aid:
            self.after_cancel(self._gen_aid)
            self._gen_aid = None
        def _tick() -> None:
            self._gen_ht = max(0.0, min(1.0,
                                        self._gen_ht + direction * 0.2))
            self._draw_gen_btn()
            if 0.0 < self._gen_ht < 1.0:
                self._gen_aid = self.after(14, _tick)
        _tick()

    def _draw_gen_btn(self, _hover: bool = False) -> None:
        cv = self._gen_cv
        cv.delete("all")
        uploaded = sum(1 for v in (self.finger_var, self.draft_var,
                                   self.ptvn_var) if v.get())
        if self._generating:
            fill = _lerp_color(PRIMARY_LITE, "#C7D2FE", 0.6)
            fg, txt = "#4338CA", "Processing..."
        elif uploaded > 0:
            fill = _lerp_color(PRIMARY, PRIMARY_HOVER, self._gen_ht)
            fg, txt = PRIMARY_FG, "Run Report"
        else:
            fill, fg, txt = GEN_OFF_BG, GEN_OFF_FG, "Run Report"
        rrect(cv, 1, 1, 167, 45, 13, fill=fill, outline="")
        cv.create_text(84, 23, text=txt, fill=fg, font=(FONT, 11, "bold"))

    # Back-compat wrappers used by _refresh / _start / _poll
    def _btn_hover(self, on: bool) -> None:
        self._draw_gen_btn(on)

    def _btn_draw(self, hover: bool = False, disabled: bool = False) -> None:
        self._draw_gen_btn(hover)

    def _file_mgmt_card(self, parent: tk.Frame) -> None:
        card = Card(parent, pad=(22,16,22,16))
        card.pack(fill="x")
        b = card.body
        lf = tk.Frame(b, bg=CARD)
        lf.pack(side="left", fill="x", expand=True)
        tk.Label(lf, text="File Management", bg=CARD, fg=TEXT,
                 font=(FONT,12,"bold")).pack(anchor="w")
        tk.Label(lf, text="Access uploaded sources and generated outputs",
                 bg=CARD, fg=MUTED, font=(FONT,9)).pack(anchor="w")
        bf = tk.Frame(b, bg=CARD)
        bf.pack(side="right")
        OutlineBtn(bf, "Open Uploaded Files",
                   self._open_folder, w=200, h=40).pack(side="left", padx=(0,8))
        OutlineBtn(bf, "Open Generated Files",
                   self._open_file, w=188, h=40).pack(side="left")

    def _summary_panel(self, parent: tk.Frame) -> None:
        card = Card(parent, pad=(22,20,22,20))
        card.pack(fill="x", pady=(0,14))
        b = card.body
        tk.Label(b, text="SUMMARY", bg=CARD, fg=MUTED,
                 font=(FONT,9,"bold")).pack(anchor="w", pady=(0,16))
        SummaryRow(b, "period", "Period",
                   self._sum_period, PRIMARY_LITE, PRIMARY).pack(fill="x",
                   pady=(0,10))
        SummaryRow(b, "days", "Working Days",
                   self._sum_wdays, AMBER_BG, AMBER).pack(fill="x",
                   pady=(0,10))
        SummaryRow(b, "files", "Files Uploaded",
                   self._sum_files, EMERALD_BOX, EMERALD).pack(fill="x")

    def _checklist_panel(self, parent: tk.Frame) -> None:
        card = Card(parent, pad=(22,20,22,20))
        card.pack(fill="x")
        b = card.body
        tk.Label(b, text="PROCESS CHECKLIST", bg=CARD, fg=MUTED,
                 font=(FONT,9,"bold")).pack(anchor="w", pady=(0,16))
        self._checks = [
            CheckItem(b, "Upload source files",        "No files uploaded"),
            CheckItem(b, "Configure report period",    ""),
            CheckItem(b, "Set working days",           ""),
            CheckItem(b, "Generate attendance report", "Pending"),
        ]
        for ch in self._checks:
            ch.pack(fill="x", pady=(0,14))

    def _footer(self) -> None:
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")
        ft = tk.Frame(self, bg=CARD, height=40)
        ft.pack(fill="x")
        ft.pack_propagate(False)
        tk.Label(ft,
                 text="All files are processed locally. No data is transmitted to external servers.",
                 bg=CARD, fg=MUTED, font=(FONT,9)
                 ).pack(side="left", padx=32, anchor="w")

    # ── Refresh ───────────────────────────────────────────────────
    def _refresh(self) -> None:
        m  = self.month_var.get()
        yr = int(self.year_var.get())
        wd = int(self.wdays_var.get())
        up = sum(1 for v in (self.finger_var, self.draft_var, self.ptvn_var)
                 if v.get())

        self._period_lbl.configure(text=f"{m}  {yr}")
        self._sum_period.set(f"{m} {yr}")
        self._sum_wdays.set(f"{wd} days")
        self._sum_files.set(f"{up} / 3 files")
        self._upload_ctr.set(f"{up}/3 uploaded")

        for key, var in (("finger", self.finger_var),
                         ("draft",  self.draft_var),
                         ("ptvn",   self.ptvn_var)):
            self._fcards[key].set_file(
                Path(var.get()).name if var.get() else None)

        sub0 = (f"{up} file{'s' if up!=1 else ''} ready"
                if up else "No files uploaded")
        self._checks[0].update("Upload source files", sub0, up > 0)
        self._checks[1].update("Configure report period",
                               f"{m} {yr}", True)
        self._checks[2].update("Set working days",
                               f"{wd} days selected", True)
        gen_sub = ("Completed" if self._generated
                   else "In progress\u2026" if self._generating else "Pending")
        self._checks[3].update("Generate attendance report", gen_sub,
                               self._generated, self._generating)

        if self._generated:
            self._gen_sub.configure(
                text=f"Completed at {self._gen_time} \u00b7 Ready to download")
        elif self._generating:
            self._gen_sub.configure(text="Processing attendance records\u2026")
        elif up == 0:
            self._gen_sub.configure(
                text="Upload at least one file to continue")
        else:
            self._gen_sub.configure(
                text=f"{up} file{'s' if up!=1 else ''} ready \u00b7 Click to generate")

        self._btn_draw()
        self._update_badge()

    def _set_progress(self, pct: float) -> None:
        self._progress = max(0.0, min(100.0, pct))
        self._pct_lbl.configure(text=f"{int(self._progress)}%")
        self._draw_progress()

    def _draw_progress(self) -> None:
        cv = self._prog_cv
        w = cv.winfo_width()
        if w < 4:
            return
        cv.delete("all")
        rrect(cv, 0, 1, w, 9, 5, fill=TRACK, outline="")
        fw = int(w * self._progress / 100)
        if fw >= 10:
            col = EMERALD if self._generated else PRIMARY
            rrect(cv, 0, 1, fw, 9, 5, fill=col, outline="")
        elif fw > 0:
            col = EMERALD if self._generated else PRIMARY
            cv.create_oval(0, 1, 9, 9, fill=col, outline="")

    def _draw_badge(self) -> None:
        cv  = self._badge_cv
        bw  = cv.winfo_width()
        if bw < 4:
            bw = int(cv["width"])
        cv.delete("all")
        txt, bg, fg = self._badge_state
        tag_rect(cv, 1, 1, bw-1, 31, 28, fill=bg, outline="")
        cv.create_text(bw//2, 16, text=txt, fill=fg,
                       font=(FONT, 9, "bold"))

    def _update_badge(self) -> None:
        if self._generated:
            self._badge_state = ("\u25cf  Report Ready", EMERALD_BG, EMERALD_DARK)
            self._badge_cv.configure(width=132, height=33)
        elif self._generating:
            self._badge_state = ("\u25cf  Processing\u2026", PRIMARY_LITE, PRIMARY)
            self._badge_cv.configure(width=120, height=33)
        else:
            self._badge_state = ("\u25cf  Idle", MUTED_BG, MUTED)
            self._badge_cv.configure(width=90, height=33)
        self._draw_badge()

    # ── Traces ────────────────────────────────────────────────────
    def _trace(self) -> None:
        self.month_var.trace_add("write", lambda *_: self._period_changed())
        self.year_var.trace_add("write",  lambda *_: self._period_changed())
        self.wdays_var.trace_add("write", lambda *_: self._refresh())

    def _period_changed(self) -> None:
        try:
            self.wdays_var.set(
                count_weekdays(int(self.year_var.get()),
                               MONTH_BY_NAME.get(self.month_var.get(),
                                                 date.today().month)))
        except Exception:
            pass
        self._refresh()

    def _sel_month(self) -> int:
        return MONTH_BY_NAME.get(self.month_var.get(), date.today().month)

    # ── File pickers ──────────────────────────────────────────────
    def _pick_finger(self) -> None:
        p = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm")])
        if p:
            self.finger_var.set(p)
            self._refresh()

    def _pick_draft(self) -> None:
        p = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm")])
        if p:
            self.draft_var.set(p)
            self._refresh()

    def _pick_ptvn(self) -> None:
        p = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm")])
        if p:
            self.ptvn_var.set(p)
            self._refresh()

    # ── Generate ──────────────────────────────────────────────────
    def _start(self) -> None:
        if self._generating:
            return
        up = sum(1 for v in (self.finger_var, self.draft_var, self.ptvn_var)
                 if v.get())
        if up == 0:
            messagebox.showwarning("No files",
                                   "Please upload at least one file.")
            return
        missing = [lbl for lbl, var in (
            ("Report Finger Scan", self.finger_var),
            ("Draft Attendance",   self.draft_var),
            ("Attendance Report",  self.ptvn_var),
        ) if not var.get()]
        if missing:
            messagebox.showwarning("Missing files",
                                   "Please select: " + ", ".join(missing))
            return

        self._generating = True
        self._generated  = False
        self._set_progress(5)
        self._status_lbl.configure(text="Starting\u2026")
        self._update_badge()
        self._btn_draw(disabled=True)
        self._refresh()
        self.thread = threading.Thread(target=self._worker, daemon=True)
        self.thread.start()

    def _worker(self) -> None:
        try:
            opts = ProcessingOptions(
                month=self._sel_month(),
                year=int(self.year_var.get()),
                fuzzy_threshold=80,
                working_days=int(self.wdays_var.get()),
            )
            out  = self._out_default()
            pout = self._ptvn_default()
            result = process_attendance_report(
                self.finger_var.get(),
                self.draft_var.get(),
                out, opts,
                ptvn_report_path=self.ptvn_var.get(),
                ptvn_output_path=pout,
                log=self.log_q.put,
            )
            self.log_q.put(f"DONE::{result}::{pout}")
        except Exception as ex:
            self.log_q.put(f"ERROR::{ex}")

    def _out_default(self) -> Path:
        base = Path(self.draft_var.get()) if self.draft_var.get() \
               else Path.home()
        return base.with_name(
            f"Attendance_{self.month_var.get()}{self.year_var.get()}.xlsx")

    def _ptvn_default(self) -> Path:
        p = Path(self.ptvn_var.get()) if self.ptvn_var.get() \
            else Path.home()
        ext = ".xlsm" if p.suffix.lower() == ".xlsm" else ".xlsx"
        return p.with_name(
            f"PTVN_{self.month_var.get()}{self.year_var.get()}{ext}")

    def _poll(self) -> None:
        try:
            while True:
                msg = self.log_q.get_nowait()
                if msg.startswith("DONE::"):
                    parts = msg[6:].split("::", 1)
                    self.out_path  = Path(parts[0])
                    self.ptvn_out  = Path(parts[1])
                    self._generating = False
                    self._generated  = True
                    self._gen_time = datetime.now().strftime("%H:%M")
                    self._set_progress(100)
                    self._status_lbl.configure(text="Done")
                    self._update_badge()
                    self._btn_draw()
                    self._show_output()
                    self._refresh()
                elif msg.startswith("ERROR::"):
                    self._generating = False
                    self._set_progress(0)
                    self._status_lbl.configure(text="Error")
                    self._update_badge()
                    self._btn_draw()
                    self._refresh()
                    messagebox.showerror("Error", msg[7:])
                else:
                    pct = min(max(self._progress + 13, 13), 92)
                    self._set_progress(pct)
                    self._status_lbl.configure(text=msg[:60])
                    self._refresh()
        except queue.Empty:
            pass
        self.after(150, self._poll)

    def _show_output(self) -> None:
        for w in self._out_frame.winfo_children():
            w.destroy()
        if not self.out_path:
            return
        row = tk.Frame(self._out_frame, bg=EMERALD_BG)
        row.pack(fill="x")
        # canvas dot
        dot = tk.Canvas(row, width=8, height=8, bg=EMERALD_BG,
                        highlightthickness=0)
        dot.pack(side="left", padx=(14,10), pady=14)
        dot.create_oval(0,0,8,8, fill=EMERALD, outline="")
        tf = tk.Frame(row, bg=EMERALD_BG)
        tf.pack(side="left", fill="x", expand=True, pady=10)
        tk.Label(tf, text=self.out_path.name,
                 bg=EMERALD_BG, fg=EMERALD_DARK,
                 font=(FONT,9,"bold"), anchor="w").pack(fill="x")
        if self.ptvn_out:
            tk.Label(tf, text=self.ptvn_out.name,
                     bg=EMERALD_BG, fg=EMERALD_DARK,
                     font=(FONT,9), anchor="w").pack(fill="x")
        OutlineBtn(row, "Open", self._open_file,
                   w=76, h=30).pack(side="right", padx=12, pady=10)

    # ── File management ───────────────────────────────────────────
    def _open_file(self) -> None:
        paths = [p for p in (self.out_path, self.ptvn_out)
                 if p and p.exists()]
        if paths:
            for p in paths:
                subprocess.run(["open", str(p)], check=False)
        else:
            messagebox.showinfo("No files",
                                "No generated files available yet.")

    def _open_folder(self) -> None:
        t = next((p.parent for p in (self.out_path, self.ptvn_out)
                  if p and p.exists()), None)
        if t:
            subprocess.run(["open", str(t)], check=False)
        elif self.finger_var.get():
            subprocess.run(
                ["open", str(Path(self.finger_var.get()).parent)],
                check=False)
        else:
            messagebox.showinfo("No folder", "No path available yet.")


def main() -> None:
    app = AttendanceReportApp()
    app.mainloop()


if __name__ == "__main__":
    main()

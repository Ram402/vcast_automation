"""
VectorCAST — Holographic Automotive Launcher
=============================================
Redesigned UI with:
  • Holographic laser HUD theme  (#0a1628 deep space navy + #00d4ff cyan glow)
  • Side-scrolling panel navigator (no tabs — circular arc menu)
  • Speedometer-style animated progress arcs for build status
  • Animated scan-line header with particle drift
  • All original functionality preserved:
      Panel 1 – Unit Test       : vcast_auto_compile3.py
      Panel 2 – Batch UT        : vcast_batch_compile.py
      Panel 3 – Integration Test: vcast_it_manual_compilation.py
      Panel 4 – Import Excel    : push modules to Batch
      Panel 5 – Build Log       : run history + live output

Usage:
    python vcast_launcher.py

Requires:
    Python 3.8+   (tkinter ships with official Windows installer)
    openpyxl      (optional, for Excel import — pip install openpyxl)
"""

import os, sys, math, time, threading, tempfile, subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime
from dataclasses import dataclass
from typing import Optional, Callable

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ══════════════════════════════════════════════════════════════════════════════
#  HOLOGRAPHIC COLOUR PALETTE  (inspired by HUD laser projection)
# ══════════════════════════════════════════════════════════════════════════════
C_VOID      = "#050d1a"     # deep space — window background
C_PANEL     = "#08142a"     # panel background
C_SURFACE   = "#0d1f3c"     # input / card surface
C_SURFACE2  = "#0a1830"     # slightly lighter surface
C_CYAN      = "#00d4ff"     # primary holographic cyan
C_CYAN_DIM  = "#006b80"     # muted cyan for borders
C_CYAN_GLOW = "#40e8ff"     # bright tip of arc
C_BLUE      = "#1a6fff"     # accent blue
C_BLUE_DIM  = "#0d3a80"     # dim blue
C_GREEN     = "#00ff9d"     # PASS / success
C_GREEN_DIM = "#007a4d"
C_AMBER     = "#ffb300"     # WARNING
C_RED       = "#ff3355"     # ERROR / FAIL
C_RED_DIM   = "#7a1a28"
C_WHITE     = "#e8f4ff"     # primary text
C_GREY      = "#4a6a8a"     # secondary text
C_BORDER    = "#0f2a4a"     # subtle divider
C_MENU_SEL  = "#0f2e52"     # selected nav item bg
C_NAV_BG    = "#060f1e"     # nav rail background

FONT_HUD    = ("Consolas", 9)
FONT_HUD_B  = ("Consolas", 9, "bold")
FONT_HUD_L  = ("Consolas", 11, "bold")
FONT_UI     = ("Segoe UI", 9)
FONT_UI_B   = ("Segoe UI", 9, "bold")
FONT_TITLE  = ("Segoe UI", 13, "bold")
FONT_BIG    = ("Segoe UI", 20, "bold")
FONT_LABEL  = ("Segoe UI", 9)
FONT_SMALL  = ("Segoe UI", 8)

# ══════════════════════════════════════════════════════════════════════════════
#  LOG MANAGER  (unchanged business logic)
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class RunRecord:
    run_id: int
    started_at: datetime
    finished_at: Optional[datetime] = None
    tab_name: str = ""
    script_name: str = ""
    exit_code: Optional[int] = None
    output: str = ""

    @property
    def duration_sec(self) -> Optional[float]:
        if self.finished_at and self.started_at:
            return (self.finished_at - self.started_at).total_seconds()
        return None

    @property
    def status(self) -> str:
        if self.exit_code is None:   return "Running"
        return "PASS" if self.exit_code == 0 else "FAIL"

    def duration_str(self) -> str:
        d = self.duration_sec
        if d is None:  return "—"
        if d < 60:     return f"{d:.1f}s"
        m, s = divmod(int(d), 60)
        if m < 60:     return f"{m}m {s}s"
        h, m = divmod(m, 60)
        return f"{h}h {m}m {s}s"

class LogManager:
    def __init__(self):
        self._records: list[RunRecord] = []
        self._next_id = 1
        self._listeners: list[Callable] = []

    def subscribe(self, cb):   self._listeners.append(cb)
    def _notify(self):
        for cb in self._listeners:
            try: cb()
            except: pass

    def start_run(self, tab_name, script_path) -> RunRecord:
        rec = RunRecord(run_id=self._next_id, started_at=datetime.now(),
                        tab_name=tab_name, script_name=os.path.basename(script_path))
        self._next_id += 1
        self._records.append(rec)
        self._notify()
        return rec

    def finish_run(self, rec, exit_code, output):
        rec.finished_at = datetime.now()
        rec.exit_code = exit_code
        rec.output = output
        self._notify()

    def append_output(self, rec, text):
        rec.output += text

    @property
    def records(self):  return list(self._records)
    def clear(self):    self._records.clear(); self._notify()

    def stats(self):
        fin = [r for r in self._records if r.finished_at]
        passed = sum(1 for r in fin if r.exit_code == 0)
        failed = sum(1 for r in fin if r.exit_code not in (None, 0))
        return {"total": len(fin), "passed": passed, "failed": failed,
                "total_time": sum(r.duration_sec or 0 for r in fin)}


# ══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _browse_dir(var):
    d = filedialog.askdirectory()
    if d: var.set(d)

def _browse_file(var, ft=None):
    f = filedialog.askopenfilename(
        filetypes=ft or [("Python files","*.py"),("All","*.*")])
    if f: var.set(f)

def _guess_script(name):
    here = os.path.dirname(os.path.abspath(__file__))
    cand = os.path.join(here, name)
    return cand if os.path.isfile(cand) else name

def _safe_remove(p):
    try: os.remove(p)
    except: pass

def _parse_modules_from_excel(path):
    if not HAS_OPENPYXL:
        raise ImportError("openpyxl required: pip install openpyxl")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows: return []
    def cs(v): return "" if v is None else str(v).strip()
    hdr = [cs(c).lower() for c in rows[0]]
    si = next((i for i,h in enumerate(hdr)
               if any(k in h for k in ("uut","stem","module","name"))
               and "file" not in h and ".c" not in h), None)
    ci = next((i for i,h in enumerate(hdr)
               if any(k in h for k in (".c","cfile","source","filename","file"))), None)
    data = rows[1:] if (si is not None or ci is not None) else rows
    si = si or 0; ci = ci or 1
    mods = []
    for row in data:
        if not row: continue
        cells = list(row)+[None,None]
        stem = cs(cells[si]); cfile = cs(cells[ci])
        if not stem and not cfile: continue
        if cfile and not cfile.lower().endswith(".c"): cfile += ".c"
        if stem and cfile: mods.append((stem, cfile))
    return mods


# ══════════════════════════════════════════════════════════════════════════════
#  SPEEDOMETER CANVAS WIDGET
# ══════════════════════════════════════════════════════════════════════════════

class Speedometer(tk.Canvas):
    """
    Arc gauge — like a car speedometer.
    value 0..100, arc sweeps from -210° to +30° (240° total).
    """
    def __init__(self, parent, size=120, label="", **kw):
        kw.update(bg=C_PANEL, highlightthickness=0,
                  width=size, height=size)
        super().__init__(parent, **kw)
        self.size   = size
        self.label  = label
        self._val   = 0
        self._anim_val = 0.0
        self._target   = 0.0
        self._status   = "idle"   # idle / running / pass / fail
        self._draw()
        self._tick()

    # ── arc maths ──────────────────────────────────────────────────────
    @staticmethod
    def _arc_colour(val, status):
        if status == "pass":    return C_GREEN
        if status == "fail":    return C_RED
        if status == "running": return C_CYAN
        if val > 80:  return C_GREEN
        if val > 40:  return C_AMBER
        return C_CYAN_DIM

    def _draw(self):
        self.delete("all")
        s  = self.size
        cx = s / 2
        cy = s / 2
        r  = s * 0.40
        pad = s * 0.10

        # ── background ring ───────────────────────────────────────────
        self.create_arc(pad, pad, s-pad, s-pad,
                        start=-210, extent=240,
                        style="arc", outline=C_BORDER, width=3)

        # ── tick marks ────────────────────────────────────────────────
        for pct in range(0, 101, 10):
            ang = math.radians(-210 + pct * 2.4)
            r_out = r + s*0.07
            r_in  = r + s*0.02
            x1 = cx + r_out * math.cos(ang)
            y1 = cy - r_out * math.sin(ang)
            x2 = cx + r_in  * math.cos(ang)
            y2 = cy - r_in  * math.sin(ang)
            self.create_line(x1,y1,x2,y2, fill=C_GREY, width=1)

        # ── filled arc ────────────────────────────────────────────────
        extent = self._anim_val * 2.4
        colour = self._arc_colour(self._anim_val, self._status)
        if extent > 0.5:
            self.create_arc(pad, pad, s-pad, s-pad,
                            start=-210, extent=extent,
                            style="arc", outline=colour, width=4)
            # glow tip
            tip_ang = math.radians(-210 + extent)
            tx = cx + r * math.cos(tip_ang)
            ty = cy - r * math.sin(tip_ang)
            glow_r = s * 0.04
            self.create_oval(tx-glow_r, ty-glow_r, tx+glow_r, ty+glow_r,
                             fill=C_CYAN_GLOW, outline="")

        # ── centre value ──────────────────────────────────────────────
        if self._status == "pass":
            val_txt  = "PASS"
            val_clr  = C_GREEN
        elif self._status == "fail":
            val_txt  = "FAIL"
            val_clr  = C_RED
        elif self._status == "running":
            val_txt  = f"{int(self._anim_val)}%"
            val_clr  = C_CYAN
        else:
            val_txt  = "IDLE"
            val_clr  = C_GREY
        self.create_text(cx, cy-4, text=val_txt,
                         fill=val_clr, font=FONT_HUD_B, anchor="center")
        if self.label:
            self.create_text(cx, cy+10, text=self.label,
                             fill=C_GREY, font=FONT_SMALL, anchor="center")

    def set_value(self, val, status="running"):
        self._target = max(0, min(100, val))
        self._status = status

    def set_idle(self):
        self._target = 0
        self._status = "idle"
        self._anim_val = 0

    def set_pass(self):
        self._target = 100
        self._status = "pass"

    def set_fail(self, pct=0):
        self._target = pct
        self._status = "fail"

    def _tick(self):
        dirty = False
        if abs(self._anim_val - self._target) > 0.5:
            self._anim_val += (self._target - self._anim_val) * 0.12
            dirty = True
        elif self._anim_val != self._target:
            self._anim_val = self._target
            dirty = True
        if dirty:
            self._draw()
        self.after(40, self._tick)


# ══════════════════════════════════════════════════════════════════════════════
#  HUD SCAN-LINE HEADER
# ══════════════════════════════════════════════════════════════════════════════

class HUDHeader(tk.Canvas):
    def __init__(self, parent, **kw):
        kw.update(bg=C_VOID, highlightthickness=0, height=72)
        super().__init__(parent, **kw)
        self._scan_y  = 0
        self._pulse   = 0.0
        self._dp      = 0.04
        self.bind("<Configure>", self._on_resize)
        self._draw()
        self._animate()

    def _on_resize(self, e):  self._draw()

    def _draw(self):
        self.delete("all")
        w = self.winfo_width() or 900
        h = 72

        # bottom glow gradient line
        self.create_line(0, h-1, w, h-1, fill=C_CYAN_DIM, width=1)
        self.create_line(0, h-2, w, h-2, fill=C_CYAN, width=1)

        # scan line
        scan_clr = _alpha_cyan(0.15 + 0.10 * math.sin(self._pulse * 2))
        self.create_line(0, self._scan_y, w, self._scan_y,
                         fill=scan_clr, width=1)

        # left bracket + title
        self.create_text(20, 36, text="[  VECTORCAST  ]",
                         fill=C_CYAN, font=("Consolas", 16, "bold"), anchor="w")
        self.create_text(220, 42, text="AUTOMOTIVE  SOFTWARE  VERIFICATION  SUITE",
                         fill=C_GREY, font=("Consolas", 8), anchor="w")

        # right block — datetime
        now = datetime.now().strftime("%Y-%m-%d   %H:%M:%S")
        self.create_text(w-16, 28, text=now,
                         fill=C_CYAN_DIM, font=FONT_HUD, anchor="e")
        self.create_text(w-16, 46, text="SYS ▸ ONLINE",
                         fill=C_GREEN, font=FONT_HUD, anchor="e")

        # decorative corner brackets
        b = 12
        for x, anchor_x in ((0, 1), (w, -1)):
            self.create_line(x, 4, x+anchor_x*b, 4, fill=C_CYAN, width=2)
            self.create_line(x, 4, x, 4+b,        fill=C_CYAN, width=2)
            self.create_line(x, h-4, x+anchor_x*b, h-4, fill=C_CYAN, width=2)
            self.create_line(x, h-4, x, h-b-4,     fill=C_CYAN, width=2)

    def _animate(self):
        w = self.winfo_width() or 900
        self._scan_y = (self._scan_y + 1) % 72
        self._pulse  = (self._pulse + self._dp) % (2 * math.pi)
        self._draw()
        self.after(35, self._animate)


def _alpha_cyan(alpha):
    # Blend C_CYAN (#00d4ff) toward C_VOID (#050d1a) by alpha
    r = int(0x00 * alpha + 0x05 * (1-alpha))
    g = int(0xd4 * alpha + 0x0d * (1-alpha))
    b = int(0xff * alpha + 0x1a * (1-alpha))
    return f"#{r:02x}{g:02x}{b:02x}"


# ══════════════════════════════════════════════════════════════════════════════
#  SIDE NAV RAIL  (replaces notebook tabs)
# ══════════════════════════════════════════════════════════════════════════════

class NavRail(tk.Frame):
    """Vertical icon + label navigation bar — automotive cockpit style."""

    ITEMS = [
        ("UT",    "UNIT\nTEST",    C_CYAN),
        ("BATCH", "BATCH\nUT",     C_BLUE),
        ("IT",    "INTEG\nTEST",   C_GREEN),
        ("XLS",   "EXCEL\nIMPORT", C_AMBER),
        ("LOG",   "BUILD\nLOGS",   C_GREY),
    ]

    def __init__(self, parent, on_select: Callable, **kw):
        kw.update(bg=C_NAV_BG, width=72)
        super().__init__(parent, **kw)
        self.pack_propagate(False)
        self._selected = 0
        self._btns: list[tk.Canvas] = []
        self._on_select = on_select
        self._build()

    def _build(self):
        # top logo mark
        logo = tk.Canvas(self, bg=C_NAV_BG, width=72, height=56,
                         highlightthickness=0)
        logo.pack(pady=(8, 0))
        logo.create_text(36, 28, text="VC", fill=C_CYAN,
                         font=("Consolas", 18, "bold"))
        logo.create_line(8, 50, 64, 50, fill=C_CYAN_DIM, width=1)

        for i, (code, label, colour) in enumerate(self.ITEMS):
            c = tk.Canvas(self, bg=C_NAV_BG, width=72, height=68,
                          highlightthickness=0, cursor="hand2")
            c.pack(pady=2)
            c.bind("<Button-1>", lambda e, idx=i: self._click(idx))
            c.bind("<Enter>",    lambda e, idx=i: self._hover(idx, True))
            c.bind("<Leave>",    lambda e, idx=i: self._hover(idx, False))
            self._btns.append(c)
        self._render_all()

    def _click(self, idx):
        self._selected = idx
        self._render_all()
        self._on_select(idx)

    def _hover(self, idx, entering):
        if idx == self._selected: return
        self._draw_item(idx, hovering=entering)

    def _render_all(self):
        for i in range(len(self.ITEMS)):
            self._draw_item(i)

    def _draw_item(self, i, hovering=False):
        c = self._btns[i]
        code, label, colour = self.ITEMS[i]
        sel = (i == self._selected)
        c.delete("all")
        bg = C_MENU_SEL if sel else ("#0b1a30" if hovering else C_NAV_BG)
        c.create_rectangle(0, 0, 72, 68, fill=bg, outline="")

        # left accent bar for selected
        if sel:
            c.create_rectangle(0, 0, 3, 68, fill=colour, outline="")

        # icon text
        icon_clr = colour if sel else (C_WHITE if hovering else C_GREY)
        c.create_text(36, 22, text=code, fill=icon_clr,
                      font=("Consolas", 10, "bold"), anchor="center")
        # label lines
        lbl_clr = C_WHITE if sel else (C_GREY if not hovering else colour)
        c.create_text(36, 48, text=label, fill=lbl_clr,
                      font=("Segoe UI", 7), anchor="center", justify="center")

        # bottom divider
        c.create_line(8, 66, 64, 66, fill=C_BORDER, width=1)

    def select(self, idx):
        self._click(idx)


# ══════════════════════════════════════════════════════════════════════════════
#  HUD FIELD ROW  (label + glowing entry + optional browse)
# ══════════════════════════════════════════════════════════════════════════════

def hud_field(parent, label, var, row, browse_fn=None, col=0, colspan=1):
    """Place a label + styled entry on a grid."""
    tk.Label(parent, text=label, bg=C_PANEL, fg=C_GREY,
             font=FONT_UI, anchor="w").grid(
        row=row, column=col, sticky="w", padx=(16, 6), pady=3)

    e = tk.Entry(parent, textvariable=var,
                 bg=C_SURFACE, fg=C_WHITE,
                 insertbackground=C_CYAN,
                 relief="flat", font=FONT_HUD, bd=0,
                 highlightthickness=1,
                 highlightbackground=C_CYAN_DIM,
                 highlightcolor=C_CYAN)
    e.grid(row=row, column=col+1, sticky="ew",
           padx=(0, 4 if browse_fn else 16), pady=3)

    if browse_fn:
        tk.Button(parent, text="…", command=browse_fn,
                  bg=C_SURFACE2, fg=C_CYAN,
                  activebackground=C_SURFACE, activeforeground=C_CYAN_GLOW,
                  relief="flat", font=FONT_HUD_B,
                  cursor="hand2", width=3, bd=0).grid(
            row=row, column=col+2, padx=(0, 16), pady=3)
    return e


def hud_section(parent, title, row, colour=C_CYAN, col=0, colspan=3):
    """Horizontal section label with glowing underline."""
    f = tk.Frame(parent, bg=colour, height=1)
    f.grid(row=row, column=col, columnspan=colspan,
           sticky="ew", padx=16, pady=(12, 0))
    tk.Label(parent, text=f"◈  {title.upper()}",
             bg=C_PANEL, fg=colour,
             font=("Consolas", 8, "bold")).grid(
        row=row+1, column=col, columnspan=colspan,
        sticky="w", padx=16, pady=(2, 4))
    return row + 2


def hud_button(parent, text, command, colour=C_CYAN, fg=C_VOID, **kw):
    return tk.Button(parent, text=text, command=command,
                     bg=colour, fg=fg,
                     activebackground=C_CYAN_GLOW,
                     activeforeground=C_VOID,
                     relief="flat", font=FONT_UI_B,
                     cursor="hand2", bd=0,
                     padx=kw.pop("padx", 18),
                     pady=kw.pop("pady", 7), **kw)


# ══════════════════════════════════════════════════════════════════════════════
#  SCROLLABLE PANEL  (replaces ttk.Notebook page)
# ══════════════════════════════════════════════════════════════════════════════

def make_scroll_panel(parent):
    """Return (frame, inner_grid_frame) with vertical scroll."""
    outer = tk.Frame(parent, bg=C_VOID)
    outer.pack(fill="both", expand=True)

    canvas = tk.Canvas(outer, bg=C_VOID, highlightthickness=0)
    sb = tk.Scrollbar(outer, orient="vertical", command=canvas.yview,
                      bg=C_NAV_BG, troughcolor=C_VOID,
                      activebackground=C_CYAN_DIM, width=8, bd=0)
    canvas.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    inner = tk.Frame(canvas, bg=C_PANEL)
    win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _cfg(_e=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(win_id, width=canvas.winfo_width())

    inner.bind("<Configure>", _cfg)
    canvas.bind("<Configure>", _cfg)

    def _scroll(e):
        if canvas.winfo_ismapped():
            canvas.yview_scroll(-1*(e.delta//120), "units")
    inner.bind_all("<MouseWheel>", _scroll)

    inner.columnconfigure(1, weight=1)
    return outer, inner


# ══════════════════════════════════════════════════════════════════════════════
#  HOLOGRAPHIC CONSOLE
# ══════════════════════════════════════════════════════════════════════════════

class HoloConsole(tk.Frame):
    def __init__(self, parent, height=12, **kw):
        kw.update(bg=C_PANEL)
        super().__init__(parent, **kw)

        # top bar
        bar = tk.Frame(self, bg=C_SURFACE2)
        bar.pack(fill="x")
        tk.Label(bar, text="◈  LIVE OUTPUT", bg=C_SURFACE2,
                 fg=C_CYAN, font=("Consolas", 8, "bold")).pack(side="left", padx=10, pady=4)
        tk.Button(bar, text="CLR", command=self.clear,
                  bg=C_SURFACE2, fg=C_GREY, relief="flat",
                  font=FONT_SMALL, cursor="hand2", bd=0).pack(side="right", padx=8)

        # left glow bar
        glow = tk.Frame(self, bg=C_CYAN_DIM, width=2)
        glow.pack(side="left", fill="y")

        # text widget
        self._txt = tk.Text(self, height=height,
                            bg="#020a14", fg=C_WHITE,
                            font=FONT_HUD, relief="flat",
                            state="disabled", wrap="word",
                            insertbackground=C_CYAN,
                            selectbackground=C_BLUE_DIM,
                            bd=0,
                            highlightthickness=1,
                            highlightbackground=C_CYAN_DIM)
        sb = tk.Scrollbar(self, command=self._txt.yview,
                          bg=C_SURFACE2, troughcolor=C_VOID, width=8, bd=0)
        self._txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._txt.pack(fill="both", expand=True)

        self._txt.tag_config("success", foreground=C_GREEN)
        self._txt.tag_config("error",   foreground=C_RED)
        self._txt.tag_config("info",    foreground=C_CYAN)
        self._txt.tag_config("warn",    foreground=C_AMBER)
        self._txt.tag_config("dim",     foreground=C_GREY)

    def write(self, text, tag=""):
        self._txt.config(state="normal")
        if tag: self._txt.insert("end", text, tag)
        else:   self._txt.insert("end", text)
        self._txt.see("end")
        self._txt.config(state="disabled")

    def clear(self):
        self._txt.config(state="normal")
        self._txt.delete("1.0", "end")
        self._txt.config(state="disabled")

    def get_all_text(self):
        return self._txt.get("1.0", "end-1c")


# ══════════════════════════════════════════════════════════════════════════════
#  RUNNER ENGINE  (unchanged logic, added speedometer hooks)
# ══════════════════════════════════════════════════════════════════════════════

def _colour_line(line):
    lo = line.lower()
    if any(k in lo for k in ("error","failed","fail","[fail")): return "error"
    if any(k in lo for k in ("success","pass","[ok]","done")):  return "success"
    if any(k in lo for k in ("step","attempt","batch","[module","compil")): return "info"
    return ""

def _run_script(script_path, extra_env, console: HoloConsole,
                run_btn: tk.Button, log_manager: LogManager,
                tab_name="Run", done_callback=None,
                clear_console=True, speedo: Speedometer = None):

    def _worker():
        run_btn.config(state="disabled", text="◉  Running…")
        if clear_console: console.clear()
        if speedo: speedo.set_value(5, "running")

        rec  = log_manager.start_run(tab_name, script_path)
        t0   = time.perf_counter()
        ts   = rec.started_at.strftime("%Y-%m-%d %H:%M:%S")
        hdr  = f"{'─'*60}\n[{ts}]  {tab_name}  ▸  {rec.script_name}\n{'─'*60}\n\n"
        console.write(hdr, "info")
        log_manager.append_output(rec, hdr)

        env  = os.environ.copy()
        env.update(extra_env)
        collected = [hdr]
        exit_code = 1
        line_count = 0

        try:
            proc = subprocess.Popen(
                [sys.executable, script_path],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, env=env,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform=="win32" else 0,
            )
            for line in proc.stdout:
                console.write(line, _colour_line(line))
                collected.append(line)
                log_manager.append_output(rec, line)
                line_count += 1
                if speedo:
                    pct = min(95, 5 + line_count * 0.5)
                    speedo.set_value(pct, "running")
            proc.wait()
            exit_code = proc.returncode
            elapsed = time.perf_counter() - t0
            log_manager.finish_run(rec, exit_code, "".join(collected))
            dur = rec.duration_str() if rec.duration_sec else f"{elapsed:.1f}s"
            tag = "success" if exit_code == 0 else "error"
            footer = (f"\n{'─'*60}\n"
                      f"Completed in {dur}  |  Exit: {exit_code}  |  "
                      f"{'PASS ✓' if exit_code==0 else 'FAIL ✗'}\n")
            console.write(footer, tag)
            log_manager.append_output(rec, footer)
            if speedo:
                if exit_code == 0: speedo.set_pass()
                else:              speedo.set_fail(int(min(95, 5+line_count*0.5)))
        except FileNotFoundError:
            msg = f"\n[ERROR] Script not found:\n  {script_path}\n"
            console.write(msg, "error")
            log_manager.finish_run(rec, 1, "".join(collected)+msg)
            if speedo: speedo.set_fail(10)
        except Exception as exc:
            msg = f"\n[ERROR] {exc}\n"
            console.write(msg, "error")
            log_manager.finish_run(rec, 1, "".join(collected)+msg)
            if speedo: speedo.set_fail(10)
        finally:
            run_btn.config(state="normal", text=run_btn._orig_text)
            if done_callback: done_callback()

    run_btn._orig_text = run_btn.cget("text")
    threading.Thread(target=_worker, daemon=True).start()


def _patch_and_run(script_path, patches, console, run_btn,
                   log_manager, tab_name, speedo=None):
    if not script_path or not os.path.isfile(script_path):
        messagebox.showerror("Script not found",
            f"Cannot find:\n{script_path}\n\nUse the … button to locate it.")
        return
    lines = [
        "import importlib.util, sys, os",
        f"_spec = importlib.util.spec_from_file_location('_mod', r'''{script_path}''')",
        "_mod = importlib.util.module_from_spec(_spec)",
        "_spec.loader.exec_module(_mod)",
    ]
    for k, v in patches.items():
        if isinstance(v, str):   lines.append(f"_mod.{k} = r'''{v}'''")
        elif isinstance(v, list):
            items = ", ".join(f'r"""{i}"""' for i in v)
            lines.append(f"_mod.{k} = [{items}]")
        elif isinstance(v, int): lines.append(f"_mod.{k} = {v}")
        else:                    lines.append(f"_mod.{k} = {repr(v)}")
    lines.append("_mod.main()")
    fd, wp = tempfile.mkstemp(suffix=".py", prefix="vcast_wrap_")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        _run_script(wp, {}, console, run_btn, log_manager, tab_name,
                    done_callback=lambda: _safe_remove(wp),
                    speedo=speedo)
    except Exception as exc:
        _safe_remove(wp)
        messagebox.showerror("Error", str(exc))


# ══════════════════════════════════════════════════════════════════════════════
#  PANEL TITLE  (holographic panel header)
# ══════════════════════════════════════════════════════════════════════════════

def panel_title(parent, title, subtitle, colour=C_CYAN, speedo_label="BUILD"):
    """Top row of a panel: title + subtitle + speedometer."""
    hdr = tk.Frame(parent, bg=C_PANEL)
    hdr.grid(row=0, column=0, columnspan=3, sticky="ew", padx=16, pady=(16, 4))
    hdr.columnconfigure(0, weight=1)

    # left: title
    left = tk.Frame(hdr, bg=C_PANEL)
    left.pack(side="left", fill="x", expand=True)
    tk.Label(left, text=title, bg=C_PANEL, fg=colour,
             font=("Consolas", 14, "bold")).pack(anchor="w")
    tk.Label(left, text=subtitle, bg=C_PANEL, fg=C_GREY,
             font=FONT_SMALL).pack(anchor="w")

    # right: speedometer
    speedo = Speedometer(hdr, size=90, label=speedo_label)
    speedo.pack(side="right", padx=(0, 8))

    # divider
    div = tk.Canvas(parent, bg=C_PANEL, height=2,
                    highlightthickness=0)
    div.grid(row=1, column=0, columnspan=3, sticky="ew", padx=16)
    # gradient line drawn after widget appears
    def _draw_div(e=None, c=div, col=colour):
        w = c.winfo_width() or 800
        c.delete("all")
        c.create_line(0, 1, w, 1, fill=col, width=1)
        c.create_line(0, 2, w//3, 2, fill=col, width=1)
    div.bind("<Configure>", _draw_div)

    return speedo


# ══════════════════════════════════════════════════════════════════════════════
#  DYNAMIC FILE LIST  (UUT / SBF rows with + / × buttons)
# ══════════════════════════════════════════════════════════════════════════════

class FileList(tk.Frame):
    """Add/remove list for .c filenames (IT UUTs and SBFs)."""
    def __init__(self, parent, colour=C_CYAN, **kw):
        kw.update(bg=C_PANEL)
        super().__init__(parent, **kw)
        self._colour = colour
        self._rows: list[tk.StringVar] = []
        self._body = tk.Frame(self, bg=C_SURFACE,
                              highlightthickness=1,
                              highlightbackground=colour)
        self._body.pack(fill="x")
        self._body.columnconfigure(0, weight=1)
        self._add_row()

    def _add_row(self, val=""):
        sv  = tk.StringVar(value=val)
        idx = len(self._rows)
        row = tk.Frame(self._body, bg=C_SURFACE)
        row.grid(row=idx, column=0, sticky="ew", padx=4, pady=2)
        row.columnconfigure(0, weight=1)
        tk.Entry(row, textvariable=sv,
                 bg=C_VOID, fg=C_WHITE,
                 insertbackground=C_CYAN,
                 relief="flat", font=FONT_HUD, bd=0,
                 highlightthickness=1,
                 highlightbackground=self._colour,
                 highlightcolor=self._colour).pack(
            side="left", fill="x", expand=True, padx=(0, 4))
        def _del(sv=sv, row=row):
            row.destroy()
            if sv in self._rows: self._rows.remove(sv)
        tk.Button(row, text="×", command=_del,
                  bg=C_SURFACE, fg=C_RED, relief="flat",
                  font=FONT_HUD_B, cursor="hand2", bd=0, width=2).pack(side="right")
        self._rows.append(sv)

    def add_btn(self, parent, label):
        """Return a button that adds a row when clicked."""
        return hud_button(parent, label, self._add_row,
                          colour=self._colour, fg=C_VOID, padx=12, pady=4)

    def get(self) -> list[str]:
        return [sv.get().strip() for sv in self._rows if sv.get().strip()]


# ══════════════════════════════════════════════════════════════════════════════
#  PANEL 1 — UNIT TEST
# ══════════════════════════════════════════════════════════════════════════════

def build_panel_ut(container, log_manager):
    _, inner = make_scroll_panel(container)
    speedo   = panel_title(inner,
                           "◈  UNIT TEST COMPILER",
                           "Single UUT  ·  auto-retry  ·  header / macro fix",
                           C_CYAN, "UT BUILD")

    v = {k: tk.StringVar(value=d) for k, d in [
        ("script",   _guess_script("vcast_auto_compile3.py")),
        ("vcast",    r"C:\VCAST"),
        ("env",      ""), ("work",   ""), ("bname",  "R2"),
        ("bpath",    ""), ("hroot",  ""), ("src1",   ""),
        ("src2",     ""), ("src3",   ""), ("uut",    ""),
        ("defines",  "__USE_MINGW_ANSI_STDIO"),
        ("inc1",""), ("inc2",""), ("inc3",""), ("retry","100"),
    ]}

    r = 2
    r = hud_section(inner, "Script",            r, C_CYAN)
    hud_field(inner,"Script path",    v["script"], r,lambda:_browse_file(v["script"])); r+=1
    r = hud_section(inner, "VectorCAST",        r, C_CYAN)
    hud_field(inner,"VectorCAST dir", v["vcast"],  r,lambda:_browse_dir(v["vcast"]));  r+=1
    r = hud_section(inner, "Environment",       r, C_CYAN)
    hud_field(inner,"ENV_NAME",       v["env"],    r); r+=1
    hud_field(inner,"WORK_DIR",       v["work"],   r,lambda:_browse_dir(v["work"]));   r+=1
    r = hud_section(inner, "Project / Base",    r, C_CYAN)
    hud_field(inner,"BASE_DIR_NAME",  v["bname"],  r); r+=1
    hud_field(inner,"BASE_DIR_PATH",  v["bpath"],  r,lambda:_browse_dir(v["bpath"]));  r+=1
    hud_field(inner,"HEADER_SEARCH",  v["hroot"],  r,lambda:_browse_dir(v["hroot"]));  r+=1
    r = hud_section(inner, "Source Directories", r, C_CYAN)
    hud_field(inner,"SOURCE_DIR_1",   v["src1"],   r,lambda:_browse_dir(v["src1"]));   r+=1
    hud_field(inner,"SOURCE_DIR_2",   v["src2"],   r,lambda:_browse_dir(v["src2"]));   r+=1
    hud_field(inner,"SOURCE_DIR_3",   v["src3"],   r,lambda:_browse_dir(v["src3"]));   r+=1
    r = hud_section(inner, "Unit Under Test",   r, C_CYAN)
    hud_field(inner,"UUT_FILE (stem)",v["uut"],    r); r+=1
    r = hud_section(inner, "Compiler / Defines", r, C_CYAN)
    hud_field(inner,"DEFINES",        v["defines"],r); r+=1
    hud_field(inner,"EXTRA_INCLUDE_1",v["inc1"],   r,lambda:_browse_dir(v["inc1"]));   r+=1
    hud_field(inner,"EXTRA_INCLUDE_2",v["inc2"],   r,lambda:_browse_dir(v["inc2"]));   r+=1
    hud_field(inner,"EXTRA_INCLUDE_3",v["inc3"],   r,lambda:_browse_dir(v["inc3"]));   r+=1
    hud_field(inner,"MAX_RETRY",      v["retry"],  r); r+=1

    r = hud_section(inner, "Live Output", r, C_CYAN)
    console = HoloConsole(inner, height=14)
    console.grid(row=r, column=0, columnspan=3, sticky="nsew", padx=16, pady=4)
    inner.rowconfigure(r, weight=1); r+=1

    def _run():
        bp = v["bpath"].get().strip()
        _patch_and_run(v["script"].get().strip(), dict(
            VECTORCAST_DIR=v["vcast"].get().strip(),
            ENV_NAME=v["env"].get().strip(),
            WORK_DIR=v["work"].get().strip(),
            BASE_DIR_NAME=v["bname"].get().strip(),
            BASE_DIR_PATH=bp,
            SOURCE_DIR_1=v["src1"].get().strip(),
            SOURCE_DIR_2=v["src2"].get().strip(),
            SOURCE_DIR_3=v["src3"].get().strip(),
            UUT_FILE=v["uut"].get().strip(),
            HEADER_SEARCH_ROOT=v["hroot"].get().strip() or bp,
            MAX_RETRY_ROUNDS=int(v["retry"].get().strip() or 100),
            DEFINES=[d.strip() for d in v["defines"].get().split() if d.strip()],
            EXTRA_INCLUDE_1=v["inc1"].get().strip(),
            EXTRA_INCLUDE_2=v["inc2"].get().strip(),
            EXTRA_INCLUDE_3=v["inc3"].get().strip(),
        ), console, run_btn, log_manager, "Unit Test", speedo=speedo)

    run_btn = hud_button(inner, "▶  INITIATE COMPILATION", _run, C_CYAN)
    run_btn.grid(row=r, column=0, columnspan=3, pady=(12, 20))


# ══════════════════════════════════════════════════════════════════════════════
#  PANEL 2 — BATCH UT
# ══════════════════════════════════════════════════════════════════════════════

def build_panel_batch(container, log_manager, app_state):
    _, inner = make_scroll_panel(container)
    speedo   = panel_title(inner,
                           "◈  BATCH UNIT TEST",
                           "Sequential multi-module compilation",
                           C_BLUE, "BATCH")

    v = {k: tk.StringVar(value=d) for k, d in [
        ("script",  _guess_script("vcast_batch_compile.py")),
        ("vcast",   r"C:\VCAST"), ("bname","R2"),
        ("bpath",""), ("hroot",""), ("workspace",""), ("retry","100"),
    ]}

    r = 2
    r = hud_section(inner, "Script",           r, C_BLUE)
    hud_field(inner,"Batch script",v["script"],r,lambda:_browse_file(v["script"]));r+=1
    r = hud_section(inner, "VectorCAST",       r, C_BLUE)
    hud_field(inner,"VectorCAST dir",v["vcast"],r,lambda:_browse_dir(v["vcast"])); r+=1
    r = hud_section(inner, "Project / Base",   r, C_BLUE)
    hud_field(inner,"BASE_DIR_NAME", v["bname"],r); r+=1
    hud_field(inner,"BASE_DIR_PATH", v["bpath"],r,lambda:_browse_dir(v["bpath"])); r+=1
    hud_field(inner,"HEADER_SEARCH", v["hroot"],r,lambda:_browse_dir(v["hroot"])); r+=1
    hud_field(inner,"WORKSPACE_ROOT",v["workspace"],r,lambda:_browse_dir(v["workspace"])); r+=1
    hud_field(inner,"MAX_RETRY",     v["retry"],r); r+=1

    # ── module table ──────────────────────────────────────────────────
    r = hud_section(inner, "Modules  (UUT stem  ↔  .c filename)", r, C_BLUE)
    tk.Label(inner, text="Add manually or push from Excel Import panel",
             bg=C_PANEL, fg=C_GREY, font=FONT_SMALL).grid(
        row=r, column=0, columnspan=3, sticky="w", padx=16); r+=1

    tbl_outer = tk.Frame(inner, bg=C_SURFACE,
                         highlightthickness=1, highlightbackground=C_BLUE)
    tbl_outer.grid(row=r, column=0, columnspan=3, sticky="ew", padx=16, pady=4)
    tbl_outer.columnconfigure(0, weight=1)
    tbl_outer.columnconfigure(1, weight=1)
    r += 1

    # header row
    hrow = tk.Frame(tbl_outer, bg=C_SURFACE2)
    hrow.grid(row=0, column=0, columnspan=3, sticky="ew")
    hrow.columnconfigure(0, weight=1); hrow.columnconfigure(1, weight=1)
    tk.Label(hrow, text="UUT STEM", bg=C_SURFACE2, fg=C_BLUE,
             font=("Consolas", 8, "bold")).grid(row=0, column=0, sticky="w", padx=8, pady=4)
    tk.Label(hrow, text=".c FILENAME", bg=C_SURFACE2, fg=C_BLUE,
             font=("Consolas", 8, "bold")).grid(row=0, column=1, sticky="w", padx=8, pady=4)

    module_rows: list = []
    body = tk.Frame(tbl_outer, bg=C_SURFACE)
    body.grid(row=1, column=0, columnspan=3, sticky="ew")
    body.columnconfigure(0, weight=1); body.columnconfigure(1, weight=1)

    def _add_mod_row(stem="", cfile=""):
        ri = len(module_rows)
        sv = tk.StringVar(value=stem); cv = tk.StringVar(value=cfile)
        bg = C_VOID if ri % 2 == 0 else C_SURFACE2
        e1 = tk.Entry(body, textvariable=sv, bg=bg, fg=C_WHITE,
                      insertbackground=C_BLUE, relief="flat",
                      font=FONT_HUD, highlightthickness=0)
        e1.grid(row=ri, column=0, sticky="ew", padx=6, pady=2)
        e2 = tk.Entry(body, textvariable=cv, bg=bg, fg=C_WHITE,
                      insertbackground=C_BLUE, relief="flat",
                      font=FONT_HUD, highlightthickness=0)
        e2.grid(row=ri, column=1, sticky="ew", padx=6, pady=2)
        def _del(sv=sv,cv=cv,e1=e1,e2=e2,b=None):
            e1.destroy(); e2.destroy()
            if (sv,cv) in module_rows: module_rows.remove((sv,cv))
        btn = tk.Button(body, text="×", command=_del,
                        bg=bg, fg=C_RED, relief="flat",
                        font=FONT_HUD_B, cursor="hand2", bd=0, width=2)
        btn.grid(row=ri, column=2, padx=4)
        module_rows.append((sv, cv))

    def _clear_mods():
        for w in body.winfo_children(): w.destroy()
        module_rows.clear()

    def _set_mods(mods, replace=True):
        if replace: _clear_mods()
        for s, c in mods: _add_mod_row(s, c)
        if not module_rows: _add_mod_row()

    _add_mod_row(); _add_mod_row()
    app_state["batch_set_modules"] = _set_mods
    app_state["batch_get_modules"] = lambda: [
        (sv.get().strip(), cv.get().strip())
        for sv, cv in module_rows
        if sv.get().strip() and cv.get().strip()
    ]

    btn_r = tk.Frame(inner, bg=C_PANEL)
    btn_r.grid(row=r, column=0, columnspan=3, pady=4)
    hud_button(btn_r, "＋  ADD MODULE", _add_mod_row,
               colour=C_BLUE, padx=12, pady=4).pack(side="left", padx=16)
    r += 1

    r = hud_section(inner, "Live Output", r, C_BLUE)
    console = HoloConsole(inner, height=12)
    console.grid(row=r, column=0, columnspan=3, sticky="nsew", padx=16, pady=4)
    inner.rowconfigure(r, weight=1); r += 1

    def _run():
        bp   = v["bpath"].get().strip()
        mods = app_state["batch_get_modules"]()
        if not mods:
            messagebox.showwarning("No Modules", "Add at least one module row."); return
        sc = v["script"].get().strip()
        if not sc or not os.path.isfile(sc):
            messagebox.showerror("Not Found", f"Script not found:\n{sc}"); return
        mod_repr = repr(mods)
        lines = [
            "import importlib.util, sys, os",
            f"_spec = importlib.util.spec_from_file_location('_mod', r'''{sc}''')",
            "_mod = importlib.util.module_from_spec(_spec)",
            "_spec.loader.exec_module(_mod)",
            f"_mod.VECTORCAST_DIR     = r'''{v['vcast'].get().strip()}'''",
            f"_mod.BASE_DIR_NAME      = r'''{v['bname'].get().strip()}'''",
            f"_mod.BASE_DIR_PATH      = r'''{bp}'''",
            f"_mod.HEADER_SEARCH_ROOT = r'''{v['hroot'].get().strip() or bp}'''",
            f"_mod.WORKSPACE_ROOT     = r'''{v['workspace'].get().strip()}'''",
            f"_mod.MAX_RETRY_ROUNDS   = {int(v['retry'].get().strip() or 100)}",
            f"_mod.MODULES            = {mod_repr}",
            "_mod.main()",
        ]
        fd, wp = tempfile.mkstemp(suffix=".py", prefix="vcast_batch_")
        with os.fdopen(fd, "w", encoding="utf-8") as f: f.write("\n".join(lines))
        _run_script(wp, {}, console, run_btn, log_manager, "Batch UT",
                    done_callback=lambda: _safe_remove(wp), speedo=speedo)

    run_btn = hud_button(inner, "▶  INITIATE BATCH COMPILATION", _run, C_BLUE)
    run_btn.grid(row=r, column=0, columnspan=3, pady=(12, 20))


# ══════════════════════════════════════════════════════════════════════════════
#  PANEL 3 — INTEGRATION TEST
# ══════════════════════════════════════════════════════════════════════════════

def build_panel_it(container, log_manager):
    _, inner = make_scroll_panel(container)
    speedo   = panel_title(inner,
                           "◈  INTEGRATION TEST",
                           "Multi-UUT build  ·  manual UUT/SBF lists",
                           C_GREEN, "IT BUILD")

    v = {k: tk.StringVar(value=d) for k, d in [
        ("script",  _guess_script("vcast_it_manual_compilation.py")),
        ("vcast",   r"C:\VCAST"), ("env",""), ("work",""),
        ("bname","R"), ("bpath",""), ("hroot",""), ("retry","100"),
        ("defines","__USE_MINGW_ANSI_STDIO"),
        ("inc1",""), ("inc2",""), ("inc3",""),
        ("stubs","__DI __EI"),
    ]}

    r = 2
    r = hud_section(inner, "Script",           r, C_GREEN)
    hud_field(inner,"IT script path", v["script"],r,lambda:_browse_file(v["script"]));r+=1
    r = hud_section(inner, "VectorCAST",       r, C_GREEN)
    hud_field(inner,"VectorCAST dir", v["vcast"], r,lambda:_browse_dir(v["vcast"])); r+=1
    r = hud_section(inner, "Environment",      r, C_GREEN)
    hud_field(inner,"ENV_NAME",       v["env"],   r); r+=1
    hud_field(inner,"WORK_DIR",       v["work"],  r,lambda:_browse_dir(v["work"]));  r+=1
    r = hud_section(inner, "Project / Base",   r, C_GREEN)
    hud_field(inner,"BASE_DIR_NAME",  v["bname"], r); r+=1
    hud_field(inner,"BASE_DIR_PATH",  v["bpath"], r,lambda:_browse_dir(v["bpath"])); r+=1
    hud_field(inner,"HEADER_SEARCH",  v["hroot"], r,lambda:_browse_dir(v["hroot"])); r+=1
    r = hud_section(inner, "Compiler / Defines", r, C_GREEN)
    hud_field(inner,"DEFINES",        v["defines"],r); r+=1
    hud_field(inner,"EXTRA_INCLUDE_1",v["inc1"],  r,lambda:_browse_dir(v["inc1"]));  r+=1
    hud_field(inner,"EXTRA_INCLUDE_2",v["inc2"],  r,lambda:_browse_dir(v["inc2"]));  r+=1
    hud_field(inner,"EXTRA_INCLUDE_3",v["inc3"],  r,lambda:_browse_dir(v["inc3"]));  r+=1
    hud_field(inner,"MAX_RETRY",      v["retry"],  r); r+=1
    hud_field(inner,"ADDITIONAL_STUBS",v["stubs"],r); r+=1

    # ── UUT list ──────────────────────────────────────────────────────
    r = hud_section(inner, "IT_UUTS  (.c filenames)", r, C_GREEN)
    uut_list = FileList(inner, colour=C_GREEN)
    uut_list.grid(row=r, column=0, columnspan=3, sticky="ew", padx=16, pady=4); r+=1
    uut_list.add_btn(inner, "＋  ADD UUT").grid(
        row=r, column=0, columnspan=3, pady=4); r+=1

    # ── SBF list ──────────────────────────────────────────────────────
    r = hud_section(inner, "IT_SBFS  (Stub-By-Function .c filenames)", r, C_AMBER)
    sbf_list = FileList(inner, colour=C_AMBER)
    sbf_list.grid(row=r, column=0, columnspan=3, sticky="ew", padx=16, pady=4); r+=1
    sbf_list.add_btn(inner, "＋  ADD SBF").grid(
        row=r, column=0, columnspan=3, pady=4); r+=1

    # ── console ───────────────────────────────────────────────────────
    r = hud_section(inner, "Live Output", r, C_GREEN)
    console = HoloConsole(inner, height=12)
    console.grid(row=r, column=0, columnspan=3, sticky="nsew", padx=16, pady=4)
    inner.rowconfigure(r, weight=1); r+=1

    def _run():
        sc = v["script"].get().strip()
        if not sc or not os.path.isfile(sc):
            messagebox.showerror("Not Found", f"Script not found:\n{sc}"); return
        bp = v["bpath"].get().strip()
        it_uuts = uut_list.get()
        it_sbfs = sbf_list.get()
        if not it_uuts:
            messagebox.showwarning("No UUTs", "Add at least one UUT entry."); return
        stubs = [s.strip() for s in v["stubs"].get().split() if s.strip()]
        lines = [
            "import importlib.util, sys, os",
            f"_spec = importlib.util.spec_from_file_location('_mod', r'''{sc}''')",
            "_mod = importlib.util.module_from_spec(_spec)",
            "_spec.loader.exec_module(_mod)",
            f"_mod.VECTORCAST_DIR     = r'''{v['vcast'].get().strip()}'''",
            f"_mod.ENV_NAME           = r'''{v['env'].get().strip()}'''",
            f"_mod.WORK_DIR           = r'''{v['work'].get().strip()}'''",
            f"_mod.BASE_DIR_NAME      = r'''{v['bname'].get().strip()}'''",
            f"_mod.BASE_DIR_PATH      = r'''{bp}'''",
            f"_mod.HEADER_SEARCH_ROOT = r'''{v['hroot'].get().strip() or bp}'''",
            f"_mod.MAX_RETRY_ROUNDS   = {int(v['retry'].get().strip() or 100)}",
            f"_mod.DEFINES            = {[d.strip() for d in v['defines'].get().split() if d.strip()]!r}",
            f"_mod.EXTRA_INCLUDE_1    = r'''{v['inc1'].get().strip()}'''",
            f"_mod.EXTRA_INCLUDE_2    = r'''{v['inc2'].get().strip()}'''",
            f"_mod.EXTRA_INCLUDE_3    = r'''{v['inc3'].get().strip()}'''",
            f"_mod.ADDITIONAL_STUBS   = {stubs!r}",
            f"_mod.IT_UUTS            = {it_uuts!r}",
            f"_mod.IT_SBFS            = {it_sbfs!r}",
            "_mod.main()",
        ]
        fd, wp = tempfile.mkstemp(suffix=".py", prefix="vcast_it_")
        with os.fdopen(fd, "w", encoding="utf-8") as f: f.write("\n".join(lines))
        _run_script(wp, {}, console, run_btn, log_manager, "Integration Test",
                    done_callback=lambda: _safe_remove(wp), speedo=speedo)

    run_btn = hud_button(inner, "▶  INITIATE IT BUILD", _run, C_GREEN)
    run_btn.grid(row=r, column=0, columnspan=3, pady=(12, 20))


# ══════════════════════════════════════════════════════════════════════════════
#  PANEL 4 — EXCEL IMPORT
# ══════════════════════════════════════════════════════════════════════════════

def build_panel_excel(container, app_state):
    _, inner = make_scroll_panel(container)
    panel_title(inner,
                "◈  EXCEL IMPORT",
                "Load module list from spreadsheet  ·  push to Batch UT",
                C_AMBER, "IMPORT")

    v_path    = tk.StringVar()
    v_replace = tk.BooleanVar(value=True)
    preview:  list = []

    r = 2
    r = hud_section(inner, "Spreadsheet", r, C_AMBER)
    hud_field(inner, "Excel file (.xlsx)", v_path, r,
              lambda: _browse_file(v_path,[("Excel","*.xlsx"),("All","*.*")])); r+=1

    hint = ("Expected: column A = UUT stem,  column B = .c filename.\n"
            "Header row auto-detected.  Without headers: A=stem, B=.c file.")
    if not HAS_OPENPYXL:
        hint += "\n⚠  openpyxl not found — run:  pip install openpyxl"
    tk.Label(inner, text=hint, bg=C_PANEL, fg=C_GREY, font=FONT_SMALL,
             justify="left").grid(row=r, column=0, columnspan=3,
                                  sticky="w", padx=16, pady=(0, 8)); r+=1

    # preview table
    r = hud_section(inner, "Preview", r, C_AMBER)
    style = ttk.Style()
    style.configure("Holo.Treeview",
                    background=C_VOID, foreground=C_WHITE,
                    fieldbackground=C_VOID, rowheight=22,
                    font=FONT_HUD, borderwidth=0)
    style.configure("Holo.Treeview.Heading",
                    background=C_SURFACE2, foreground=C_AMBER,
                    font=("Consolas", 8, "bold"), relief="flat")
    style.map("Holo.Treeview",
              background=[("selected", C_BLUE_DIM)],
              foreground=[("selected", C_CYAN)])

    tv_wrap = tk.Frame(inner, bg=C_SURFACE,
                       highlightthickness=1, highlightbackground=C_AMBER)
    tv_wrap.grid(row=r, column=0, columnspan=3, sticky="nsew", padx=16, pady=4)
    tv_wrap.columnconfigure(0, weight=1)
    tv_wrap.rowconfigure(0, weight=1)
    r += 1

    tv = ttk.Treeview(tv_wrap, columns=("s","c"), show="headings",
                      style="Holo.Treeview", height=12)
    tv.heading("s", text="UUT STEM")
    tv.heading("c", text=".c FILENAME")
    tv.column("s", width=300, anchor="w")
    tv.column("c", width=300, anchor="w")
    tsb = tk.Scrollbar(tv_wrap, command=tv.yview,
                       bg=C_SURFACE2, troughcolor=C_VOID, width=8, bd=0)
    tv.configure(yscrollcommand=tsb.set)
    tv.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
    tsb.grid(row=0, column=1, sticky="ns")

    status = tk.Label(inner, text="No file loaded",
                      bg=C_PANEL, fg=C_GREY, font=FONT_UI)
    status.grid(row=r, column=0, columnspan=3, sticky="w", padx=16, pady=4); r+=1

    def _load():
        nonlocal preview
        path = v_path.get().strip()
        if not path: messagebox.showwarning("No file","Select an Excel file first."); return
        if not os.path.isfile(path): messagebox.showerror("Not found", f"File not found:\n{path}"); return
        try:
            preview = _parse_modules_from_excel(path)
        except Exception as exc:
            messagebox.showerror("Import failed", str(exc)); return
        for item in tv.get_children(): tv.delete(item)
        for s, c in preview:  tv.insert("","end", values=(s,c))
        status.config(text=f"{len(preview)} module(s) loaded",
                      fg=C_GREEN if preview else C_GREY)
        if not preview:
            messagebox.showwarning("Empty","No valid rows found in spreadsheet.")

    def _push():
        if not preview:
            messagebox.showwarning("Nothing to import","Load an Excel file first."); return
        setter = app_state.get("batch_set_modules")
        if not setter:
            messagebox.showerror("Error","Batch panel not ready."); return
        setter(preview, replace=v_replace.get())
        messagebox.showinfo("Imported",
            f"{len(preview)} modules sent to Batch UT panel.\n"
            "Switch to BATCH UT to review and run.")

    btn_bar = tk.Frame(inner, bg=C_PANEL)
    btn_bar.grid(row=r, column=0, columnspan=3, pady=(8,14))
    hud_button(btn_bar,"⬆  LOAD EXCEL",_load, C_AMBER,padx=14).pack(side="left",padx=(16,6))
    hud_button(btn_bar,"→  SEND TO BATCH",_push,C_BLUE,padx=14).pack(side="left",padx=6)
    tk.Checkbutton(btn_bar, text="Replace existing modules",
                   variable=v_replace,
                   bg=C_PANEL, fg=C_GREY,
                   activebackground=C_PANEL, activeforeground=C_WHITE,
                   selectcolor=C_SURFACE, font=FONT_UI).pack(side="left",padx=16)


# ══════════════════════════════════════════════════════════════════════════════
#  PANEL 5 — BUILD LOGS
# ══════════════════════════════════════════════════════════════════════════════

def build_panel_log(container, log_manager):
    outer = tk.Frame(container, bg=C_VOID)
    outer.pack(fill="both", expand=True)
    outer.columnconfigure(0, weight=1)
    outer.rowconfigure(1, weight=1)

    # ── stats row with mini speedometers ──────────────────────────────
    stats_bar = tk.Frame(outer, bg=C_PANEL,
                         highlightthickness=1, highlightbackground=C_BORDER)
    stats_bar.grid(row=0, column=0, sticky="ew", padx=0, pady=0)

    stat_labels: dict = {}
    speedo_total  = Speedometer(stats_bar, size=80, label="RUNS")
    speedo_pass   = Speedometer(stats_bar, size=80, label="PASS%")
    speedo_fail   = Speedometer(stats_bar, size=80, label="FAIL%")

    speedo_total.pack(side="left", padx=(12,4), pady=6)
    speedo_pass.pack( side="left", padx=4,      pady=6)
    speedo_fail.pack( side="left", padx=4,      pady=6)

    tk.Frame(stats_bar, bg=C_BORDER, width=1).pack(side="left", fill="y", padx=8)

    for key, title, colour in (
        ("total",      "TOTAL RUNS",    C_GREY),
        ("passed",     "PASSED",        C_GREEN),
        ("failed",     "FAILED",        C_RED),
        ("total_time", "TOTAL TIME",    C_CYAN),
    ):
        cell = tk.Frame(stats_bar, bg=C_PANEL, padx=16, pady=8)
        cell.pack(side="left")
        tk.Label(cell, text=title, bg=C_PANEL, fg=C_GREY,
                 font=("Consolas", 7, "bold")).pack(anchor="w")
        lbl = tk.Label(cell, text="0", bg=C_PANEL, fg=colour,
                       font=("Consolas", 18, "bold"))
        lbl.pack(anchor="w")
        stat_labels[key] = lbl

    # ── main area: history list + output pane ─────────────────────────
    pane = tk.Frame(outer, bg=C_VOID)
    pane.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
    pane.columnconfigure(0, weight=1)
    pane.columnconfigure(1, weight=2)
    pane.rowconfigure(0, weight=1)

    # left: history tree
    left = tk.Frame(pane, bg=C_PANEL,
                    highlightthickness=1, highlightbackground=C_BORDER)
    left.grid(row=0, column=0, sticky="nsew", padx=(0,1))
    left.rowconfigure(1, weight=1)
    left.columnconfigure(0, weight=1)

    tk.Label(left, text="◈  RUN HISTORY", bg=C_PANEL, fg=C_CYAN,
             font=("Consolas", 9, "bold")).grid(row=0, column=0, sticky="w",
                                                padx=10, pady=6)

    hist_wrap = tk.Frame(left, bg=C_VOID)
    hist_wrap.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0,8))
    hist_wrap.columnconfigure(0, weight=1); hist_wrap.rowconfigure(0, weight=1)

    style = ttk.Style()
    style.configure("Log.Treeview",
                    background=C_VOID, foreground=C_WHITE,
                    fieldbackground=C_VOID, rowheight=24,
                    font=FONT_HUD, borderwidth=0)
    style.configure("Log.Treeview.Heading",
                    background=C_SURFACE2, foreground=C_CYAN,
                    font=("Consolas", 8, "bold"), relief="flat")
    style.map("Log.Treeview",
              background=[("selected", C_BLUE_DIM)],
              foreground=[("selected", C_CYAN)])

    cols = ("time","tab","dur","status")
    hist = ttk.Treeview(hist_wrap, columns=cols, show="headings",
                        style="Log.Treeview", height=24)
    hist.heading("time",   text="STARTED")
    hist.heading("tab",    text="TYPE")
    hist.heading("dur",    text="TIME")
    hist.heading("status", text="STATUS")
    hist.column("time",   width=130, anchor="w")
    hist.column("tab",    width=110, anchor="w")
    hist.column("dur",    width=70,  anchor="center")
    hist.column("status", width=60,  anchor="center")
    hsb = tk.Scrollbar(hist_wrap, command=hist.yview,
                       bg=C_SURFACE2, troughcolor=C_VOID, width=8, bd=0)
    hist.configure(yscrollcommand=hsb.set)
    hist.grid(row=0, column=0, sticky="nsew")
    hsb.grid(row=0, column=1, sticky="ns")

    btn_row = tk.Frame(left, bg=C_PANEL)
    btn_row.grid(row=2, column=0, sticky="w", padx=8, pady=(0,8))
    hud_button(btn_row,"↺ REFRESH", lambda:_refresh(),
               C_SURFACE2, fg=C_CYAN, padx=10, pady=4).pack(side="left", padx=(0,4))

    # right: output pane
    right = tk.Frame(pane, bg=C_PANEL,
                     highlightthickness=1, highlightbackground=C_BORDER)
    right.grid(row=0, column=1, sticky="nsew", padx=(1,0))
    right.rowconfigure(1, weight=1)
    right.columnconfigure(0, weight=1)

    tk.Label(right, text="◈  RUN OUTPUT", bg=C_PANEL, fg=C_CYAN,
             font=("Consolas", 9, "bold")).grid(row=0, column=0, sticky="w",
                                                padx=10, pady=6)
    detail = HoloConsole(right, height=28)
    detail.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0,8))

    record_map: dict = {}

    def _fmt_time(sec):
        if sec < 60: return f"{sec:.0f}s"
        m, s = divmod(int(sec), 60)
        return f"{m}m {s}s" if m < 60 else f"{m//60}h {m%60}m"

    def _refresh(_e=None):
        for item in hist.get_children(): hist.delete(item)
        record_map.clear()
        for rec in reversed(log_manager.records):
            iid = str(rec.run_id)
            record_map[iid] = rec
            hist.insert("", "end", iid=iid, values=(
                rec.started_at.strftime("%Y-%m-%d %H:%M:%S"),
                rec.tab_name, rec.duration_str(), rec.status,
            ))
        st = log_manager.stats()
        stat_labels["total"].config(text=str(st["total"]))
        stat_labels["passed"].config(text=str(st["passed"]))
        stat_labels["failed"].config(text=str(st["failed"]))
        stat_labels["total_time"].config(text=_fmt_time(st["total_time"]))

        total = st["total"]
        if total:
            speedo_total.set_value(min(100, total * 10), "running")
            speedo_pass.set_value(st["passed"]/total*100, "pass" if st["passed"]==total else "running")
            speedo_fail.set_value(st["failed"]/total*100, "fail" if st["failed"]>0 else "running")
        else:
            speedo_total.set_idle(); speedo_pass.set_idle(); speedo_fail.set_idle()

        ch = hist.get_children()
        if ch and not hist.selection():
            hist.selection_set(ch[0]); _show()

    def _show(_e=None):
        sel = hist.selection()
        if not sel: return
        rec = record_map.get(sel[0])
        if not rec: return
        detail.clear()
        hdr = (f"Run #{rec.run_id}  |  {rec.tab_name}  |  {rec.script_name}\n"
               f"Started: {rec.started_at.strftime('%Y-%m-%d %H:%M:%S')}  "
               f"|  Duration: {rec.duration_str()}  |  {rec.status}\n"
               f"{'═'*70}\n\n")
        tag = "success" if rec.exit_code==0 else "error" if rec.exit_code else "info"
        detail.write(hdr, tag)
        detail.write(rec.output or "(no output captured)")

    hist.bind("<<TreeviewSelect>>", _show)
    log_manager.subscribe(_refresh)
    _refresh()


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN WINDOW
# ══════════════════════════════════════════════════════════════════════════════

def main():
    root = tk.Tk()
    root.title("VectorCAST  ◈  Holographic Automotive Launcher")
    root.geometry("1180x840")
    root.minsize(900, 640)
    root.configure(bg=C_VOID)

    # ttk scrollbar default
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("TScrollbar",
                    background=C_SURFACE2, troughcolor=C_VOID,
                    arrowcolor=C_GREY, borderwidth=0)

    log_manager = LogManager()
    app_state:  dict = {}

    # ── animated HUD header ──────────────────────────────────────────
    HUDHeader(root).pack(fill="x")

    # ── body: nav rail + content area ────────────────────────────────
    body = tk.Frame(root, bg=C_VOID)
    body.pack(fill="both", expand=True)

    content = tk.Frame(body, bg=C_VOID)
    content.pack(side="left", fill="both", expand=True)

    # Build all panels (hidden initially)
    panels = []
    for builder, args in [
        (build_panel_ut,    (log_manager,)),
        (build_panel_batch, (log_manager, app_state)),
        (build_panel_it,    (log_manager,)),
        (build_panel_excel, (app_state,)),
        (build_panel_log,   (log_manager,)),
    ]:
        frame = tk.Frame(content, bg=C_VOID)
        frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        builder(frame, *args)
        panels.append(frame)

    def _switch(idx):
        for i, p in enumerate(panels):
            if i == idx: p.lift()

    # nav rail (placed after content so it overlays on left)
    nav = NavRail(body, on_select=_switch)
    nav.pack(side="left", fill="y", before=content)

    _switch(0)

    # ── status bar ───────────────────────────────────────────────────
    status = tk.Frame(root, bg=C_PANEL, height=22,
                      highlightthickness=1, highlightbackground=C_BORDER)
    status.pack(fill="x", side="bottom")
    status.pack_propagate(False)
    tk.Label(status,
             text="VectorCAST Automotive Verification Suite  ◈  Holographic Edition",
             bg=C_PANEL, fg=C_GREY, font=("Consolas", 7)).pack(side="left", padx=10)
    sys_lbl = tk.Label(status, text="● SYSTEM READY", bg=C_PANEL,
                       fg=C_GREEN, font=("Consolas", 7))
    sys_lbl.pack(side="right", padx=10)

    root.mainloop()


if __name__ == "__main__":
    main()

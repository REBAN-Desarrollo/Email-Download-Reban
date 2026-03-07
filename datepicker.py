import tkinter as tk
from tkinter import ttk
import calendar
from datetime import date

from utils import MONTHS_ES, MONTHS_IMAP, DAYS_ES


class DatePicker(tk.Frame):
    """Widget de fecha con boton que despliega un calendario popup."""

    def __init__(self, parent, initial_date=None, **kwargs):
        super().__init__(parent, **kwargs)

        if initial_date:
            self._date = initial_date
        else:
            self._date = date.today()

        self._var = tk.StringVar(value=self._format_imap(self._date))

        self.entry = ttk.Entry(self, textvariable=self._var, width=14)
        self.entry.pack(side="left")
        self.btn = ttk.Button(self, text="...", width=3, command=self._open_popup)
        self.btn.pack(side="left", padx=(2, 0))

        self._popup = None

    def get(self):
        return self._var.get().strip()

    def set_date(self, d):
        self._date = d
        self._var.set(self._format_imap(d))

    def set_text(self, text):
        self._var.set(text)
        try:
            self._date = self._parse_imap(text)
        except Exception:
            pass

    def _format_imap(self, d):
        return f"{d.day:02d}-{MONTHS_IMAP[d.month - 1]}-{d.year}"

    def _parse_imap(self, s):
        parts = s.split("-")
        day = int(parts[0])
        month = MONTHS_IMAP.index(parts[1]) + 1
        year = int(parts[2])
        return date(year, month, day)

    def _open_popup(self):
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
            self._popup = None
            return

        try:
            self._date = self._parse_imap(self._var.get().strip())
        except Exception:
            self._date = date.today()

        self._popup = _CalendarPopup(self, self._date, self._on_select)

    def _on_select(self, selected_date):
        self._date = selected_date
        self._var.set(self._format_imap(selected_date))
        if self._popup:
            self._popup.destroy()
            self._popup = None


class _CalendarPopup(tk.Toplevel):
    """Popup de calendario con navegacion mes/ano."""

    def __init__(self, parent, current_date, callback):
        super().__init__(parent)
        self.overrideredirect(True)
        self.callback = callback
        self.current_year = current_date.year
        self.current_month = current_date.month
        self.selected_date = current_date

        x = parent.winfo_rootx()
        y = parent.winfo_rooty() + parent.winfo_height() + 2
        self.geometry(f"+{x}+{y}")

        self.configure(bg="#ffffff", bd=1, relief="solid")

        # Frame de navegacion
        nav = tk.Frame(self, bg="#2c3e50")
        nav.pack(fill="x")

        btn_style = {"bg": "#2c3e50", "fg": "white", "bd": 0,
                     "font": ("Arial", 10, "bold"), "cursor": "hand2",
                     "activebackground": "#34495e", "activeforeground": "white"}

        tk.Button(nav, text="<<", command=self._prev_year, **btn_style).pack(side="left", padx=2)
        tk.Button(nav, text="<", command=self._prev_month, **btn_style).pack(side="left", padx=2)

        self.lbl_header = tk.Label(nav, text="", bg="#2c3e50", fg="white",
                                   font=("Arial", 11, "bold"), cursor="hand2")
        self.lbl_header.pack(side="left", expand=True, fill="x")
        self.lbl_header.bind("<Button-1>", self._show_month_year_selector)

        tk.Button(nav, text=">", command=self._next_month, **btn_style).pack(side="right", padx=2)
        tk.Button(nav, text=">>", command=self._next_year, **btn_style).pack(side="right", padx=2)

        # Frame para selector rapido mes/ano (oculto por defecto)
        self.quick_frame = tk.Frame(self, bg="#ecf0f1")
        self._quick_visible = False

        # Frame del calendario
        self.cal_frame = tk.Frame(self, bg="white")
        self.cal_frame.pack(padx=2, pady=2)

        # Boton "Hoy"
        tk.Button(self, text="Hoy", command=self._select_today,
                  bg="#3498db", fg="white", bd=0, cursor="hand2",
                  font=("Arial", 9), activebackground="#2980b9").pack(fill="x", padx=2, pady=(0, 2))

        self._build_calendar()

        self.bind("<FocusOut>", self._on_focus_out)
        self.focus_set()
        self.grab_set()

    def _on_focus_out(self, event):
        try:
            focused = self.focus_get()
            if focused and (focused == self or str(focused).startswith(str(self))):
                return
        except Exception:
            pass
        self.destroy()

    def _build_calendar(self):
        for w in self.cal_frame.winfo_children():
            w.destroy()

        self.lbl_header.config(
            text=f"{MONTHS_ES[self.current_month - 1]} {self.current_year}"
        )

        for i, day_name in enumerate(DAYS_ES):
            tk.Label(self.cal_frame, text=day_name, width=4, bg="#ecf0f1",
                     fg="#2c3e50", font=("Arial", 9, "bold")).grid(row=0, column=i, padx=1, pady=1)

        cal = calendar.monthcalendar(self.current_year, self.current_month)
        today = date.today()

        for row_idx, week in enumerate(cal, 1):
            for col_idx, day in enumerate(week):
                if day == 0:
                    tk.Label(self.cal_frame, text="", width=4, bg="white").grid(
                        row=row_idx, column=col_idx, padx=1, pady=1)
                else:
                    d = date(self.current_year, self.current_month, day)
                    bg = "white"
                    fg = "#2c3e50"
                    font = ("Arial", 9)

                    if d == today:
                        bg = "#3498db"
                        fg = "white"
                        font = ("Arial", 9, "bold")
                    if d == self.selected_date:
                        bg = "#e74c3c"
                        fg = "white"
                        font = ("Arial", 9, "bold")

                    btn = tk.Label(self.cal_frame, text=str(day), width=4,
                                   bg=bg, fg=fg, font=font, cursor="hand2",
                                   relief="flat")
                    btn.grid(row=row_idx, column=col_idx, padx=1, pady=1)
                    btn.bind("<Button-1>", lambda e, dd=d: self._on_day_click(dd))
                    btn.bind("<Enter>", lambda e, b=btn: b.config(bg="#bdc3c7") if b.cget("bg") == "white" else None)
                    btn.bind("<Leave>", lambda e, b=btn, bg_orig=bg: b.config(bg=bg_orig))

    def _on_day_click(self, d):
        self.callback(d)

    def _prev_month(self):
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self._build_calendar()

    def _next_month(self):
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self._build_calendar()

    def _prev_year(self):
        self.current_year -= 1
        self._build_calendar()

    def _next_year(self):
        self.current_year += 1
        self._build_calendar()

    def _select_today(self):
        self.callback(date.today())

    def _show_month_year_selector(self, event=None):
        if self._quick_visible:
            self.quick_frame.pack_forget()
            self._quick_visible = False
            return

        for w in self.quick_frame.winfo_children():
            w.destroy()

        year_frame = tk.Frame(self.quick_frame, bg="#ecf0f1")
        year_frame.pack(fill="x", padx=5, pady=(5, 2))
        tk.Label(year_frame, text="Ano:", bg="#ecf0f1",
                 font=("Arial", 9, "bold")).pack(side="left")

        year_var = tk.IntVar(value=self.current_year)
        tk.Button(year_frame, text="-", width=2, command=lambda: self._quick_change_year(year_var, -1),
                  bg="#bdc3c7", bd=0).pack(side="left", padx=2)
        self.year_lbl = tk.Label(year_frame, textvariable=year_var, bg="#ecf0f1",
                                  font=("Arial", 10, "bold"), width=6)
        self.year_lbl.pack(side="left")
        tk.Button(year_frame, text="+", width=2, command=lambda: self._quick_change_year(year_var, 1),
                  bg="#bdc3c7", bd=0).pack(side="left", padx=2)
        self._year_var = year_var

        month_frame = tk.Frame(self.quick_frame, bg="#ecf0f1")
        month_frame.pack(fill="x", padx=5, pady=(2, 5))

        for i, m_name in enumerate(MONTHS_ES):
            r, c = divmod(i, 4)
            bg = "#3498db" if (i + 1) == self.current_month else "#dfe6e9"
            fg = "white" if (i + 1) == self.current_month else "#2c3e50"
            btn = tk.Label(month_frame, text=m_name, width=5, bg=bg, fg=fg,
                           font=("Arial", 9), cursor="hand2", relief="flat")
            btn.grid(row=r, column=c, padx=2, pady=2)
            btn.bind("<Button-1>", lambda e, mi=i+1, yv=year_var: self._quick_select_month(mi, yv))

        self.quick_frame.pack(fill="x", before=self.cal_frame)
        self._quick_visible = True

    def _quick_change_year(self, year_var, delta):
        year_var.set(year_var.get() + delta)

    def _quick_select_month(self, month, year_var):
        self.current_month = month
        self.current_year = year_var.get()
        self.quick_frame.pack_forget()
        self._quick_visible = False
        self._build_calendar()

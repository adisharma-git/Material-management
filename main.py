"""
Hospital Pharmacy Inventory Management
Tabbed GUI – Scripts 1, 2, 4
"""

import sys
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ── Make sure the src package is importable when frozen ──────────────────────
if getattr(sys, "frozen", False):
    BASE_DIR = sys._MEIPASS  # PyInstaller temp dir
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, BASE_DIR)

from src.script1_global_stock    import create_global_stock_lookup
from src.script2_main_store_stock import create_main_store_stock_lookup
from src.script4_inventory_calc  import InventoryCalculator


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────
ACCENT   = "#1F4E79"
BG       = "#F5F7FA"
BTN_BG   = "#2980B9"
BTN_FG   = "white"
SUCCESS  = "#27AE60"
ERR_CLR  = "#E74C3C"
FONT_H   = ("Segoe UI", 11, "bold")
FONT_N   = ("Segoe UI", 10)
FONT_LOG = ("Courier New", 9)


def pick_file(entry: tk.Entry, title: str, filetypes):
    """Open file dialog and put path into entry."""
    path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)


def pick_save(entry: tk.Entry, title: str, default_ext: str, filetypes):
    path = filedialog.asksaveasfilename(
        title=title, defaultextension=default_ext, filetypes=filetypes
    )
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)


def labeled_row(parent, label: str, row: int, browse_cmd, is_save=False):
    """Create a label + entry + Browse button row. Returns the entry widget."""
    tk.Label(parent, text=label, font=FONT_N, bg=BG, anchor="w").grid(
        row=row, column=0, sticky="w", padx=8, pady=4
    )
    entry = tk.Entry(parent, font=FONT_N, width=52)
    entry.grid(row=row, column=1, padx=4, pady=4, sticky="ew")
    btn_text = "Save As…" if is_save else "Browse…"
    tk.Button(
        parent, text=btn_text, font=FONT_N,
        bg=BTN_BG, fg=BTN_FG, relief="flat",
        command=browse_cmd,
    ).grid(row=row, column=2, padx=4, pady=4)
    return entry


def log_widget(parent) -> tk.Text:
    """Return a scrollable log Text widget."""
    frame = tk.Frame(parent, bg=BG)
    frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    txt = tk.Text(frame, font=FONT_LOG, height=12, state="disabled",
                  bg="#1C1C1C", fg="#DCDCDC", wrap="word")
    sb  = ttk.Scrollbar(frame, command=txt.yview)
    txt.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    txt.pack(side="left", fill="both", expand=True)
    return txt


def log_append(txt: tk.Text, msg: str):
    """Thread-safe log append."""
    def _do():
        txt.configure(state="normal")
        txt.insert("end", msg + "\n")
        txt.see("end")
        txt.configure(state="disabled")
    txt.after(0, _do)


def run_in_thread(fn, on_done):
    """Run fn() in background; call on_done(err_or_None) on main thread."""
    def worker():
        try:
            fn()
            on_done(None)
        except Exception as e:
            on_done(e)
    threading.Thread(target=worker, daemon=True).start()


# ─────────────────────────────────────────────────────────────────────────────
# Tab 1 – Global Stock Lookup
# ─────────────────────────────────────────────────────────────────────────────
class Tab1(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)

        tk.Label(self, text="Script 1 · Global Stock Lookup Generator",
                 font=FONT_H, bg=BG, fg=ACCENT).pack(pady=(14, 2))
        tk.Label(self, text="Aggregates stock across ALL hospital locations from the global CSV report.",
                 font=FONT_N, bg=BG, fg="#555").pack(pady=(0, 10))

        grid = tk.Frame(self, bg=BG)
        grid.pack(fill="x", padx=10)
        grid.columnconfigure(1, weight=1)

        CSV_FT  = [("CSV files", "*.csv"), ("All files", "*.*")]
        XLSX_FT = [("Excel files", "*.xlsx"), ("All files", "*.*")]

        self.in_entry = labeled_row(
            grid, "Input CSV (Global Stock Report):", 0,
            lambda: pick_file(self.in_entry, "Select Global Stock Report CSV", CSV_FT),
        )
        self.out_entry = labeled_row(
            grid, "Output Excel (Global Lookup):", 1,
            lambda: pick_save(self.out_entry, "Save Global Stock Lookup As", ".xlsx", XLSX_FT),
            is_save=True,
        )
        self.out_entry.insert(0, "Material_Global_Stock_Lookup.xlsx")

        self._run_btn = tk.Button(
            self, text="▶  Generate Global Stock Lookup",
            font=("Segoe UI", 11, "bold"),
            bg=SUCCESS, fg="white", relief="flat", padx=20, pady=8,
            command=self._run,
        )
        self._run_btn.pack(pady=10)

        self._status = tk.Label(self, text="", font=FONT_N, bg=BG)
        self._status.pack()

        self._log = log_widget(self)

    def _run(self):
        inp  = self.in_entry.get().strip()
        outp = self.out_entry.get().strip()
        if not inp or not outp:
            messagebox.showwarning("Missing", "Please set both input and output files.")
            return

        self._run_btn.configure(state="disabled", text="Running…")
        self._status.configure(text="", fg="black")

        def task():
            create_global_stock_lookup(inp, outp, log_fn=lambda m: log_append(self._log, m))

        def done(err):
            self._run_btn.configure(state="normal", text="▶  Generate Global Stock Lookup")
            if err:
                self._status.configure(text=f"❌ Error: {err}", fg=ERR_CLR)
                log_append(self._log, f"\nERROR: {err}")
            else:
                self._status.configure(text="✅ Done! File saved.", fg=SUCCESS)

        run_in_thread(task, done)


# ─────────────────────────────────────────────────────────────────────────────
# Tab 2 – Main Store Stock Lookup
# ─────────────────────────────────────────────────────────────────────────────
class Tab2(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)

        tk.Label(self, text="Script 2 · Main Store Stock Lookup Generator",
                 font=FONT_H, bg=BG, fg=ACCENT).pack(pady=(14, 2))
        tk.Label(self, text="Aggregates batch-level stock from the Main Medical Store into item-level totals.",
                 font=FONT_N, bg=BG, fg="#555").pack(pady=(0, 10))

        grid = tk.Frame(self, bg=BG)
        grid.pack(fill="x", padx=10)
        grid.columnconfigure(1, weight=1)

        XLSX_FT = [("Excel files", "*.xlsx"), ("All files", "*.*")]

        self.in_entry = labeled_row(
            grid, "Input Excel (Stock Report without category):", 0,
            lambda: pick_file(self.in_entry, "Select Stock Report Excel", XLSX_FT),
        )
        self.out_entry = labeled_row(
            grid, "Output Excel (Main Store Lookup):", 1,
            lambda: pick_save(self.out_entry, "Save Main Store Lookup As", ".xlsx", XLSX_FT),
            is_save=True,
        )
        self.out_entry.insert(0, "Material_Main_Store_Stock_Lookup.xlsx")

        self._run_btn = tk.Button(
            self, text="▶  Generate Main Store Stock Lookup",
            font=("Segoe UI", 11, "bold"),
            bg=SUCCESS, fg="white", relief="flat", padx=20, pady=8,
            command=self._run,
        )
        self._run_btn.pack(pady=10)

        self._status = tk.Label(self, text="", font=FONT_N, bg=BG)
        self._status.pack()

        self._log = log_widget(self)

    def _run(self):
        inp  = self.in_entry.get().strip()
        outp = self.out_entry.get().strip()
        if not inp or not outp:
            messagebox.showwarning("Missing", "Please set both input and output files.")
            return

        self._run_btn.configure(state="disabled", text="Running…")
        self._status.configure(text="", fg="black")

        def task():
            create_main_store_stock_lookup(inp, outp, log_fn=lambda m: log_append(self._log, m))

        def done(err):
            self._run_btn.configure(state="normal", text="▶  Generate Main Store Stock Lookup")
            if err:
                self._status.configure(text=f"❌ Error: {err}", fg=ERR_CLR)
                log_append(self._log, f"\nERROR: {err}")
            else:
                self._status.configure(text="✅ Done! File saved.", fg=SUCCESS)

        run_in_thread(task, done)


# ─────────────────────────────────────────────────────────────────────────────
# Tab 3 – Inventory Calculation (Script 4)
# ─────────────────────────────────────────────────────────────────────────────
class Tab3(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)

        tk.Label(self, text="Script 4 · Inventory Calculation Sheet Generator",
                 font=FONT_H, bg=BG, fg=ACCENT).pack(pady=(14, 2))
        tk.Label(
            self,
            text="Combines master data, global stock, main store stock & pending POs → full reorder sheet.",
            font=FONT_N, bg=BG, fg="#555",
        ).pack(pady=(0, 10))

        grid = tk.Frame(self, bg=BG)
        grid.pack(fill="x", padx=10)
        grid.columnconfigure(1, weight=1)

        XLSX_FT = [("Excel files", "*.xlsx"), ("All files", "*.*")]
        ANY_FT  = [("Excel/CSV", "*.xlsx *.csv"), ("All files", "*.*")]

        self.master_entry = labeled_row(
            grid, "Master Data (MASTER_DATA_INPUT.xlsx):", 0,
            lambda: pick_file(self.master_entry, "Select Master Data", XLSX_FT),
        )
        self.global_entry = labeled_row(
            grid, "Global Stock Lookup:", 1,
            lambda: pick_file(self.global_entry, "Select Global Stock Lookup", XLSX_FT),
        )
        self.main_entry = labeled_row(
            grid, "Main Store Stock Lookup:", 2,
            lambda: pick_file(self.main_entry, "Select Main Store Lookup", XLSX_FT),
        )
        self.expected_entry = labeled_row(
            grid, "Expected Items (PO pending):", 3,
            lambda: pick_file(self.expected_entry, "Select Expected Items File", ANY_FT),
        )
        self.out_entry = labeled_row(
            grid, "Output Excel (Inventory Calculation):", 4,
            lambda: pick_save(self.out_entry, "Save Inventory Calculation As", ".xlsx", XLSX_FT),
            is_save=True,
        )
        self.out_entry.insert(0, "INVENTORY_CALCULATION.xlsx")

        # Tip label
        tk.Label(
            self,
            text="💡 Tip: Run Script 1 & 2 first to generate the two lookup files.",
            font=FONT_N, bg=BG, fg="#888",
        ).pack(pady=(2, 6))

        self._run_btn = tk.Button(
            self, text="▶  Generate Inventory Calculation",
            font=("Segoe UI", 11, "bold"),
            bg=SUCCESS, fg="white", relief="flat", padx=20, pady=8,
            command=self._run,
        )
        self._run_btn.pack(pady=6)

        self._status = tk.Label(self, text="", font=FONT_N, bg=BG)
        self._status.pack()

        self._log = log_widget(self)

    def _run(self):
        fields = {
            "Master Data":         self.master_entry.get().strip(),
            "Global Stock Lookup": self.global_entry.get().strip(),
            "Main Store Lookup":   self.main_entry.get().strip(),
            "Expected Items":      self.expected_entry.get().strip(),
            "Output file":         self.out_entry.get().strip(),
        }
        missing = [k for k, v in fields.items() if not v]
        if missing:
            messagebox.showwarning("Missing", f"Please set: {', '.join(missing)}")
            return

        self._run_btn.configure(state="disabled", text="Running…")
        self._status.configure(text="", fg="black")

        def task():
            calc = InventoryCalculator(log_fn=lambda m: log_append(self._log, m))
            calc.run(
                master_file=fields["Master Data"],
                global_stock_file=fields["Global Stock Lookup"],
                main_store_file=fields["Main Store Lookup"],
                expected_items_file=fields["Expected Items"],
                output_file=fields["Output file"],
            )

        def done(err):
            self._run_btn.configure(state="normal", text="▶  Generate Inventory Calculation")
            if err:
                self._status.configure(text=f"❌ Error: {err}", fg=ERR_CLR)
                log_append(self._log, f"\nERROR: {err}")
            else:
                self._status.configure(text="✅ Done! File saved.", fg=SUCCESS)

        run_in_thread(task, done)


# ─────────────────────────────────────────────────────────────────────────────
# Main App Window
# ─────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Hospital Pharmacy Inventory Manager")
        self.geometry("820x620")
        self.minsize(720, 520)
        self.configure(bg=BG)

        # Header banner
        hdr = tk.Frame(self, bg=ACCENT, height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(
            hdr,
            text="🏥  Hospital Pharmacy Inventory Manager",
            font=("Segoe UI", 14, "bold"),
            bg=ACCENT, fg="white",
        ).pack(side="left", padx=16)

        # Notebook (tabs)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook",        background=BG)
        style.configure("TNotebook.Tab",    padding=[12, 6], font=FONT_N)
        style.map("TNotebook.Tab",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", "white")])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=6, pady=6)

        nb.add(Tab1(nb), text="  Script 1 – Global Stock  ")
        nb.add(Tab2(nb), text="  Script 2 – Main Store  ")
        nb.add(Tab3(nb), text="  Script 4 – Inventory Calc  ")

        # Footer
        tk.Label(
            self,
            text="© Hospital Pharmacy · All scripts run locally – no data is sent externally.",
            font=("Segoe UI", 8), bg=BG, fg="#AAA",
        ).pack(pady=(0, 4))


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

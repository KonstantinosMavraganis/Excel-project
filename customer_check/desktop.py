import os
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .excel import export_report, read_table, sheets
from .reconcile import compare


def _plain(text):
    return "".join(c for c in unicodedata.normalize("NFD", text.casefold()) if not unicodedata.combining(c))


class InputPanel(ttk.LabelFrame):
    def __init__(self, parent, app, title, registry=False):
        super().__init__(parent, text=title, padding=12)
        self.app = app
        self.registry = registry
        self.table = None
        self.path = tk.StringVar()
        self.sheet = tk.StringVar()
        self.header = tk.StringVar(value="1")
        self.code = tk.StringVar()
        self.status = tk.StringVar()
        self.name = tk.StringVar(value="(καμία)")
        self.columnconfigure(1, weight=1)
        ttk.Button(self, text="Επιλογή Excel…", command=self.choose).grid(row=0, column=0, sticky="w")
        ttk.Label(self, textvariable=self.path, wraplength=450).grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(self, text="Φύλλο").grid(row=1, column=0, sticky="w", pady=8)
        self.sheet_box = ttk.Combobox(self, textvariable=self.sheet, state="readonly")
        self.sheet_box.grid(row=1, column=1, sticky="ew")
        self.sheet_box.bind("<<ComboboxSelected>>", self.source_changed)
        ttk.Label(self, text="Γραμμή επικεφαλίδων").grid(row=2, column=0, sticky="w")
        ttk.Spinbox(self, from_=1, to=10000, textvariable=self.header, width=7).grid(row=2, column=1, sticky="w")
        self.header.trace_add("write", lambda *_: self.source_changed())
        ttk.Button(self, text="Ανάγνωση φύλλου", command=self.load).grid(row=3, column=0, columnspan=2, sticky="w", pady=8)
        ttk.Label(self, text="Κωδικός πελάτη").grid(row=4, column=0, sticky="w")
        self.code_box = ttk.Combobox(self, textvariable=self.code, state="readonly")
        self.code_box.grid(row=4, column=1, sticky="ew")
        self.code_box.bind("<<ComboboxSelected>>", lambda _: app.invalidate())
        self.info = tk.StringVar(value="Δεν έχει διαβαστεί αρχείο.")
        if registry:
            ttk.Label(self, text="Όνομα / Επωνυμία").grid(row=5, column=0, sticky="w", pady=8)
            self.name_box = ttk.Combobox(self, textvariable=self.name, state="readonly")
            self.name_box.grid(row=5, column=1, sticky="ew")
            self.name_box.bind("<<ComboboxSelected>>", lambda _: app.invalidate())
            ttk.Label(self, text="Κατάσταση συμβολαίου").grid(row=6, column=0, sticky="w")
            self.status_box = ttk.Combobox(self, textvariable=self.status, state="readonly")
            self.status_box.grid(row=6, column=1, sticky="ew")
            self.status_box.bind("<<ComboboxSelected>>", lambda _: self.update_values())
            ttk.Label(self, text="Επίλεξε τις τιμές που σημαίνουν ενεργό:").grid(row=7, column=0, columnspan=2, sticky="w", pady=(10, 4))
            self.active = tk.Listbox(self, selectmode="multiple", exportselection=False, height=5)
            self.active.grid(row=8, column=0, columnspan=2, sticky="ew")
            self.active.bind("<<ListboxSelect>>", lambda _: app.invalidate())
        ttk.Label(self, textvariable=self.info).grid(row=9, column=0, columnspan=2, sticky="w", pady=(10, 0))

    def source_changed(self, *_):
        self.table = None
        self.info.set("Πάτησε «Ανάγνωση φύλλου» μετά την αλλαγή.")
        self.app.invalidate()

    def choose(self):
        path = filedialog.askopenfilename(title="Επιλογή Excel", filetypes=[("Excel", "*.xlsx *.xlsm"), ("Όλα τα αρχεία", "*.*")])
        if not path:
            return
        self.path.set(path)
        self.source_changed()
        # Capture the header value on the UI thread before starting background work.
        try:
            header = int(self.header.get())
        except ValueError:
            messagebox.showerror("Γραμμή επικεφαλίδων", "Δώσε ακέραιο αριθμό γραμμής.")
            return
        def loaded(names):
            self.sheet_box.configure(values=names)
            if not names:
                raise ValueError("Το αρχείο δεν περιέχει φύλλα.")
            self.sheet.set(names[0])
            self.app.submit(lambda: read_table(path, names[0], header), self.accept)
        self.app.submit(lambda: sheets(path), loaded)

    def load(self):
        try:
            path, sheet, header = self.path.get(), self.sheet.get(), int(self.header.get())
            if not path or not sheet:
                raise ValueError("Επίλεξε αρχείο και φύλλο.")
        except ValueError as exc:
            messagebox.showerror("Ανάγνωση", str(exc))
            return
        self.source_changed()
        self.app.submit(lambda: read_table(path, sheet, header), self.accept)

    def accept(self, table):
        self.table = table
        self.code_box.configure(values=table.headers)
        self.code.set(self.guess(table.headers, ["κωδικ", "customer id", "customer code", "client id"]))
        if self.registry:
            self.name_box.configure(values=["(καμία)", *table.headers])
            self.name.set(self.guess(table.headers, ["επωνυμ", "ονομα", "name"], "(καμία)"))
            self.status_box.configure(values=table.headers)
            self.status.set(self.guess(table.headers, ["συμβολ", "κατασταση", "contract", "active", "status"]))
            self.update_values()
        self.info.set(f"{len(table.rows):,} γραμμές · {len(table.headers)} στήλες")
        self.app.invalidate()

    @staticmethod
    def guess(headers, fragments, fallback=""):
        return next((h for h in headers if any(f in _plain(h) for f in fragments)), fallback)

    def index(self, variable):
        if self.table is None or variable.get() not in self.table.headers:
            raise ValueError("Διάβασε τα δύο φύλλα και επίλεξε τις απαιτούμενες στήλες.")
        return self.table.headers.index(variable.get())

    def update_values(self):
        self.active.delete(0, tk.END)
        if self.table and self.status.get() in self.table.headers:
            index = self.index(self.status)
            values = sorted({row.values[index] for row in self.table.rows if row.values[index]})
            for i, value in enumerate(values):
                self.active.insert(tk.END, value)
                if _plain(value) in {"ναι", "ενεργο", "ενεργος", "ενεργη", "yes", "true", "active", "1"}:
                    self.active.selection_set(i)
        self.app.invalidate()


class ResultsPanel(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding=8)
        self.entries = []
        self.query = tk.StringVar()
        top = ttk.Frame(self)
        top.pack(fill="x")
        ttk.Label(top, text="Αναζήτηση:").pack(side="left")
        ttk.Entry(top, textvariable=self.query).pack(side="left", padx=8, fill="x", expand=True)
        self.count = tk.StringVar()
        ttk.Label(top, textvariable=self.count).pack(side="right")
        self.tree = ttk.Treeview(self, columns=("code", "name", "row", "detail"), show="headings", height=12)
        for column, label, width in [("code", "Κωδικός", 150), ("name", "Όνομα / Επωνυμία", 220), ("row", "Γραμμή Excel", 90), ("detail", "Λεπτομέρεια", 420)]:
            self.tree.heading(column, text=label)
            self.tree.column(column, width=width, minwidth=60)
        self.tree.pack(side="left", fill="both", expand=True, pady=(8, 0))
        scroll = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y", pady=(8, 0))
        self.tree.configure(yscrollcommand=scroll.set)
        self.query.trace_add("write", lambda *_: self.render())

    def set_entries(self, entries):
        self.entries = entries
        self.query.set("")

    def render(self):
        self.tree.delete(*self.tree.get_children())
        query = self.query.get().strip().casefold()
        matched = [entry for entry in self.entries if query in f"{entry.code} {entry.name} {entry.detail}".casefold()]
        for entry in matched[:1000]:
            self.tree.insert("", "end", values=(entry.code, entry.name, entry.source_row, entry.detail))
        self.count.set(f"{min(len(matched), 1000):,} / {len(matched):,} · Η εξαγωγή περιέχει όλα")


class App:
    def __init__(self, root):
        self.root = root
        self.executor = ThreadPoolExecutor(max_workers=1)
        self.result = None
        self.metadata = {}
        self.busy = False
        root.title("Έλεγχος ενεργών πελατών")
        root.geometry("1180x850")
        root.minsize(960, 760)
        root.protocol("WM_DELETE_WINDOW", self.close)
        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("TButton", padding=7)
        outer = ttk.Frame(root, padding=16)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Έλεγχος ενεργών πελατών", font=("Segoe UI", 20, "bold")).pack(anchor="w")
        ttk.Label(outer, text="Σύγκρινε μητρώο και πληρωμές με κωδικό πελάτη. Τα αρχεία επεξεργάζονται τοπικά.").pack(anchor="w", pady=(4, 12))
        inputs = ttk.Frame(outer)
        inputs.pack(fill="x")
        inputs.columnconfigure((0, 1), weight=1, uniform="panels")
        self.registry = InputPanel(inputs, self, "1. Μητρώο πελατών", registry=True)
        self.registry.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.payments = InputPanel(inputs, self, "2. Αρχείο πληρωμών")
        self.payments.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=12)
        ttk.Button(actions, text="Έλεγχος", command=self.check).pack(side="left")
        self.export_button = ttk.Button(actions, text="Εξαγωγή αναφοράς Excel…", command=self.export, state="disabled")
        self.export_button.pack(side="left", padx=10)
        self.message = tk.StringVar(value="Επίλεξε τα δύο Excel και επιβεβαίωσε τις στήλες και τις ενεργές τιμές.")
        ttk.Label(outer, textvariable=self.message, wraplength=1100).pack(anchor="w", pady=(0, 8))
        self.tabs = ttk.Notebook(outer)
        self.tabs.pack(fill="both", expand=True)
        self.results = {}
        self.titles = {"missing": "Λείπουν", "found": "Βρέθηκαν", "unknown": "Άγνωστοι κωδικοί", "issues": "Προς έλεγχο"}
        for key, title in self.titles.items():
            panel = ResultsPanel(self.tabs)
            self.results[key] = panel
            self.tabs.add(panel, text=title)
        ttk.Label(outer, text="Ο έλεγχος αφορά παρουσία εγγραφής, όχι εξόφληση. Για στήλες με τύπους, ανανέωσε και αποθήκευσε πρώτα τα Excel.", wraplength=1100).pack(anchor="w", pady=(10, 0))

    def invalidate(self):
        self.result = None
        if hasattr(self, "export_button"):
            self.export_button.configure(state="disabled")
            self.message.set("Εκτέλεσε έλεγχο με τις τρέχουσες επιλογές.")
            for key, panel in self.results.items():
                panel.set_entries([])
                self.tabs.tab(panel, text=self.titles[key])

    def submit(self, work, success):
        if self.busy:
            return
        self.busy = True
        self.message.set("Επεξεργασία…")
        states = []
        def lock(widget):
            for child in widget.winfo_children():
                try:
                    previous = str(child.cget("state"))
                    child.configure(state="disabled")
                    states.append((child, previous))
                except tk.TclError:
                    pass
                lock(child)
        lock(self.root)
        future = self.executor.submit(work)
        def poll():
            if not future.done():
                self.root.after(80, poll)
                return
            self.busy = False
            for widget, state in states:
                widget.configure(state=state)
            try:
                self.message.set("Η επεξεργασία ολοκληρώθηκε.")
                success(future.result())
            except Exception as exc:
                self.message.set("Η ενέργεια δεν ολοκληρώθηκε. Δες το μήνυμα σφάλματος.")
                messagebox.showerror("Σφάλμα", str(exc))
        self.root.after(80, poll)

    def check(self):
        self.invalidate()
        try:
            registry, payments = self.registry, self.payments
            code, status, payment_code = registry.index(registry.code), registry.index(registry.status), payments.index(payments.code)
            name = registry.index(registry.name) if registry.name.get() != "(καμία)" else None
            active = {registry.active.get(i) for i in registry.active.curselection()}
            if not active:
                raise ValueError("Επίλεξε τουλάχιστον μία τιμή για ενεργό συμβόλαιο.")
            self.metadata = {
                "Χρόνος ελέγχου": datetime.now().astimezone().isoformat(timespec="seconds"),
                "Αρχείο μητρώου": registry.table.source,
                "Φύλλο μητρώου": registry.table.sheet,
                "Κωδικός μητρώου": registry.code.get(),
                "Στήλη συμβολαίου": registry.status.get(),
                "Ενεργές τιμές": ", ".join(sorted(active)),
                "Αρχείο πληρωμών": payments.table.source,
                "Φύλλο πληρωμών": payments.table.sheet,
                "Κωδικός πληρωμών": payments.code.get(),
            }
            registry_table, payment_table = registry.table, payments.table
        except ValueError as exc:
            messagebox.showerror("Επιλογές ελέγχου", str(exc))
            return
        self.submit(lambda: compare(registry_table, payment_table, code, status, payment_code, active, name), self.show_result)

    def show_result(self, result):
        self.result = result
        for key, panel in self.results.items():
            entries = getattr(result, key)
            panel.set_entries(entries)
            self.tabs.tab(panel, text=f"{self.titles[key]} ({len(entries):,})")
        self.export_button.configure(state="normal")
        total = len(result.found) + len(result.missing)
        self.message.set(f"Ενεργοί με μοναδικό κωδικό: {total:,} · Βρέθηκαν: {len(result.found):,} · Λείπουν: {len(result.missing):,} · Προς έλεγχο: {len(result.issues):,} · Επιπλέον εγγραφές πληρωμών: {result.duplicate_payments:,}")
        if result.issues:
            self.message.set(self.message.get() + " · Έλεγξε τις αμφίσημες / ελλιπείς εγγραφές πριν βγάλεις συμπέρασμα.")
        if total == 0:
            self.message.set(self.message.get() + " · Δεν βρέθηκαν ενεργοί πελάτες με έγκυρα, μοναδικά στοιχεία.")

    def export(self):
        if self.result is None:
            return
        path = filedialog.asksaveasfilename(title="Αποθήκευση αναφοράς", defaultextension=".xlsx", initialfile="έλεγχος_πελατών.xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        target = Path(path).resolve()
        for source in (self.registry.table.source, self.payments.table.source):
            if os.path.normcase(str(target)) == os.path.normcase(str(Path(source).resolve())) or (target.exists() and os.path.samefile(target, source)):
                messagebox.showerror("Αποθήκευση", "Επίλεξε διαφορετικό αρχείο από τα αρχικά Excel.")
                return
        result, metadata = self.result, self.metadata.copy()
        self.submit(lambda: export_report(path, result, metadata), lambda _: self.message.set(f"Αποθηκεύτηκε η αναφορά: {path}"))

    def close(self):
        self.executor.shutdown(wait=False, cancel_futures=True)
        self.root.destroy()


def run():
    root = tk.Tk()
    App(root)
    root.mainloop()

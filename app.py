# -*- coding: utf-8 -*-
"""
Bank hisobotlarini soddalashtiruvchi desktop ilova.

Bir nechta xom Hamkorbank Client-Bank hisobot fayllarini (.xlsx) tanlab,
chiqish papkasini belgilab, "Boshlash" tugmasini bosish orqali har biri
ALOHIDA-ALOHIDA soddalashtirilgan faylga (+ tekshirish ro'yxati) aylantirib
beradi. Jarayon ro'yxatda va progress bar orqali jonli kuzatiladi.
"""
import os
import sys
import subprocess
import threading
import traceback
import queue
import winreg
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk, filedialog, messagebox

import sv_ttk

import build_report
import categorize


def find_excel_exe():
    """Locate the real Excel.exe via the Windows "App Paths" registry,
    bypassing whatever program .xlsx happens to be (mis)associated with on
    this machine (a common support issue: someone once chose "Open with ->
    Notepad -> always use this app" for .xlsx files)."""
    candidates = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE"),
    ]
    for hive, subkey in candidates:
        try:
            with winreg.OpenKey(hive, subkey) as key:
                path, _ = winreg.QueryValueEx(key, "")
        except OSError:
            continue
        if path and os.path.exists(path):
            return path
    return None


def resource_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = resource_base_dir()


class FileRow:
    def __init__(self, path):
        self.path = path
        self.status = "Kutmoqda"
        self.error = None
        self.out_path = None


class UnresolvedDialog(tk.Toplevel):
    """Modal oyna: hisobotlarda kategoriyasi aniqlanmagan (nomlanmagan)
    kontragentlar ro'yxatini ko'rsatadi va har biri uchun nom/kategoriya
    kiritishni so'raydi. Kiritilgan qiymatlar persist_dictionary.json ga
    saqlanadi va keyingi barcha loyihalarda avtomatik tanilib qoladi."""

    def __init__(self, parent, unresolved):
        super().__init__(parent)
        self.title("Nomlanmagan kontragentlar topildi")
        self.geometry("760x520")
        self.transient(parent)
        self.grab_set()
        self.result_entries = {}  # key -> (info, tk.StringVar)
        self.confirmed = False

        header = ttk.Frame(self, padding=10)
        header.pack(fill="x")
        ttk.Label(
            header,
            text=(
                f"{len(unresolved)} ta kontragent hech qanday kategoriyaga to'g'ri kelmadi.\n"
                "Har biri uchun kategoriya nomini kiriting (bo'sh qoldirsangiz \"?\" bo'lib qoladi). "
                "Kiritganlaringiz keyingi loyihalar uchun ham eslab qolinadi."
            ),
            wraplength=720,
            justify="left",
        ).pack(anchor="w")

        canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=(0, 10))
        scrollbar.pack(side="left", fill="y", pady=(0, 10))

        for key, info in sorted(unresolved.items(), key=lambda kv: -kv[1]["count"]):
            row = ttk.Frame(inner, padding=6, relief="groove", borderwidth=1)
            row.pack(fill="x", padx=6, pady=4)

            label_bits = []
            if info["name"]:
                label_bits.append(info["name"])
            if info["inn"]:
                label_bits.append(f"ИНН {info['inn']}")
            label_bits.append(f"{info['count']} qatorda uchraydi")
            ttk.Label(row, text="  |  ".join(label_bits), font=("", 9, "bold")).pack(anchor="w")
            if info["sample"]:
                ttk.Label(row, text=info["sample"], foreground="#555", wraplength=680).pack(anchor="w")

            entry_row = ttk.Frame(row)
            entry_row.pack(fill="x", pady=(4, 0))
            ttk.Label(entry_row, text="Kategoriya nomi:").pack(side="left")
            var = tk.StringVar(value="")
            ttk.Entry(entry_row, textvariable=var, width=30).pack(side="left", padx=(6, 0))
            self.result_entries[key] = (info, var)

        footer = ttk.Frame(self, padding=10)
        footer.pack(fill="x")
        ttk.Button(footer, text="Saqlash va davom etish", command=self._on_confirm).pack(side="right")
        ttk.Button(footer, text="Bekor qilish", command=self._on_cancel).pack(side="right", padx=(0, 8))

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _on_confirm(self):
        self.confirmed = True
        self.destroy()

    def _on_cancel(self):
        self.confirmed = False
        self.destroy()

    def get_assignments(self):
        """Returns list of (inn, name, category) for entries the user filled in."""
        out = []
        for key, (info, var) in self.result_entries.items():
            cat = var.get().strip()
            if cat:
                out.append((info["inn"], info["name"], cat))
        return out


class GroupsManagerDialog(tk.Toplevel):
    """Guruhlar (kategoriyalar) boshqaruv oynasi. Foydalanuvchi ilovani
    birinchi marta ishga tushirganda — yoki istalgan payt — ma'lum hisob
    raqam (ИНН) yoki kompaniya nomi uchun guruh (kategoriya) belgilab
    qo'yishi mumkin, fayl tashlanishini kutmasdan. Barcha yozuvlar
    persist_dictionary.json ga saqlanadi va shu zahoti ishlatila boshlaydi."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Guruhlarni boshqarish")
        self.geometry("720x560")
        self.transient(parent)
        self.grab_set()

        self.mapping = dict(categorize.load_all_mappings())
        self.rows = {}  # key -> (row_frame, id_var, cat_var)

        header = ttk.Frame(self, padding=14)
        header.pack(fill="x")
        ttk.Label(header, text="Guruhlarni boshqarish", font=("Segoe UI Semibold", 13)).pack(anchor="w")
        ttk.Label(
            header,
            text=(
                "Bu yerda ma'lum hisob raqam (ИНН) yoki kompaniya nomi qaysi guruh (kategoriya)ga "
                "tegishli ekanini oldindan belgilab qo'yishingiz mumkin — fayl tashlanganda ilova "
                "avtomatik shu guruhga ajratadi."
            ),
            wraplength=680, foreground="#666666", justify="left",
        ).pack(anchor="w", pady=(4, 0))

        add_box = ttk.LabelFrame(self, text="Yangi guruh qo'shish", padding=10)
        add_box.pack(fill="x", padx=14, pady=(10, 0))

        self.new_type = tk.StringVar(value="inn")
        type_row = ttk.Frame(add_box)
        type_row.pack(fill="x")
        ttk.Radiobutton(type_row, text="ИНН bo'yicha", variable=self.new_type, value="inn").pack(side="left")
        ttk.Radiobutton(type_row, text="Nomi bo'yicha", variable=self.new_type, value="name").pack(side="left", padx=(12, 0))

        fields_row = ttk.Frame(add_box)
        fields_row.pack(fill="x", pady=(8, 0))
        ttk.Label(fields_row, text="ИНН / Nomi:").pack(side="left")
        self.new_id_var = tk.StringVar()
        ttk.Entry(fields_row, textvariable=self.new_id_var, width=22).pack(side="left", padx=(6, 16))
        ttk.Label(fields_row, text="Guruh nomi:").pack(side="left")
        self.new_cat_var = tk.StringVar()
        ttk.Entry(fields_row, textvariable=self.new_cat_var, width=22).pack(side="left", padx=(6, 16))
        ttk.Button(fields_row, text="+ Qo'shish", command=self._add_row).pack(side="left")

        pick_row = ttk.Frame(add_box)
        pick_row.pack(fill="x", pady=(8, 0))
        ttk.Button(pick_row, text="📄 Excel'dan tanlash...", command=self._pick_from_excel).pack(side="left")
        ttk.Label(
            pick_row,
            text="— ИНН'ni qo'lda yozish o'rniga, haqiqiy hisobot faylini ochib, kontragentlarni belgilab tanlang.",
            foreground="#888888",
        ).pack(side="left", padx=(8, 0))

        list_label = ttk.Label(self, text="Mavjud guruhlar:", font=("Segoe UI Semibold", 10))
        list_label.pack(anchor="w", padx=14, pady=(14, 4))

        canvas_frame = ttk.Frame(self)
        canvas_frame.pack(fill="both", expand=True, padx=14)
        self.canvas = tk.Canvas(canvas_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="left", fill="y")

        self._render_rows()

        footer = ttk.Frame(self, padding=14)
        footer.pack(fill="x")
        ttk.Button(footer, text="Saqlash", style="Accent.TButton", command=self._save).pack(side="right")
        ttk.Button(footer, text="Yopish", command=self.destroy).pack(side="right", padx=(0, 8))

    def _render_rows(self):
        for child in self.inner.winfo_children():
            child.destroy()
        self.rows = {}
        if not self.mapping:
            ttk.Label(self.inner, text="Hali hech qanday guruh belgilanmagan.", foreground="#888888").pack(
                anchor="w", pady=10
            )
            return
        for key, cat in sorted(self.mapping.items()):
            self._add_row_widget(key, cat)

    def _add_row_widget(self, key, category):
        is_name = key.startswith(categorize._NAME_KEY_PREFIX)
        display_id = key[len(categorize._NAME_KEY_PREFIX):] if is_name else key
        label_prefix = "Nomi: " if is_name else "ИНН: "

        row = ttk.Frame(self.inner, padding=6, relief="groove", borderwidth=1)
        row.pack(fill="x", pady=3)
        ttk.Label(row, text=f"{label_prefix}{display_id}", width=32, anchor="w").pack(side="left")
        cat_var = tk.StringVar(value=category)
        ttk.Entry(row, textvariable=cat_var, width=22).pack(side="left", padx=(6, 6))
        ttk.Button(row, text="Saqlash", command=lambda k=key, v=cat_var: self._update_row(k, v)).pack(side="left")
        ttk.Button(row, text="O'chirish", command=lambda k=key: self._delete_row(k)).pack(side="left", padx=(6, 0))
        self.rows[key] = (row, cat_var)

    def _update_row(self, key, var):
        new_cat = var.get().strip()
        if not new_cat:
            messagebox.showwarning("Diqqat", "Guruh nomi bo'sh bo'lishi mumkin emas.")
            return
        self.mapping[key] = new_cat
        categorize.save_all_mappings(self.mapping)
        messagebox.showinfo("Saqlandi", "Guruh yangilandi.")

    def _delete_row(self, key):
        if not messagebox.askyesno("Tasdiqlash", "Ushbu guruhni o'chirishni tasdiqlaysizmi?"):
            return
        self.mapping.pop(key, None)
        categorize.save_all_mappings(self.mapping)
        self._render_rows()

    def _pick_from_excel(self):
        path = filedialog.askopenfilename(
            title="Hisobot faylini tanlang",
            filetypes=[("Excel fayllar", "*.xlsx"), ("Barcha fayllar", "*.*")],
        )
        if not path:
            return
        try:
            _wb, _ws, _hdr, rows = categorize.load_raw_rows(path)
        except Exception as e:
            messagebox.showerror("Xato", f"Faylni o'qib bo'lmadi:\n{e}")
            return
        parties = categorize.unique_counterparties(rows)
        if not parties:
            messagebox.showinfo("Diqqat", "Faylda kontragentlar topilmadi.")
            return
        picker = CounterpartyPickerDialog(self, parties)
        self.wait_window(picker)
        if picker.applied:
            self.mapping = dict(categorize.load_all_mappings())
            self._render_rows()

    def _add_row(self):
        ident = self.new_id_var.get().strip()
        cat = self.new_cat_var.get().strip()
        if not ident or not cat:
            messagebox.showwarning("Diqqat", "ИНН/Nomi va Guruh nomini kiriting.")
            return
        key = ident if self.new_type.get() == "inn" else f"{categorize._NAME_KEY_PREFIX}{ident}"
        self.mapping[key] = cat
        categorize.save_all_mappings(self.mapping)
        self.new_id_var.set("")
        self.new_cat_var.set("")
        self._render_rows()

    def _save(self):
        categorize.save_all_mappings(self.mapping)
        self.destroy()


class CounterpartyPickerDialog(tk.Toplevel):
    """Xom hisobot faylini jadval (Excel'ga o'xshash) ko'rinishida ochib,
    foydalanuvchi bir nechta kontragentni belgilab (check qilib), bittasiga
    guruh nomi berib bir yo'la qo'sha oladigan oyna. INN'ni qo'lda yozish
    o'rniga, haqiqiy fayldan tanlab olish uchun."""

    def __init__(self, parent, parties):
        super().__init__(parent)
        self.title("Fayldan kontragent tanlash")
        self.geometry("860x560")
        self.transient(parent)
        self.grab_set()
        self.applied = False
        self.parties = parties
        self.checked = set()

        header = ttk.Frame(self, padding=(14, 14, 14, 6))
        header.pack(fill="x")
        ttk.Label(header, text="Fayldan kontragent tanlash", font=("Segoe UI Semibold", 13)).pack(anchor="w")
        ttk.Label(
            header,
            text="Kerakli qatorlarni belgilang (katakchani bosing), so'ng pastda guruh nomini kiritib qo'shing.",
            foreground="#666666", wraplength=820,
        ).pack(anchor="w", pady=(2, 0))

        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill="both", expand=True, padx=14, pady=(8, 0))

        columns = ("check", "inn", "name", "sample")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="none", height=16)
        self.tree.heading("check", text="✓")
        self.tree.heading("inn", text="ИНН")
        self.tree.heading("name", text="Nomi")
        self.tree.heading("sample", text="Namuna matn")
        self.tree.column("check", width=36, anchor="center")
        self.tree.column("inn", width=110, anchor="w")
        self.tree.column("name", width=220, anchor="w")
        self.tree.column("sample", width=420, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Button-1>", self._on_click)

        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scroll.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)

        for key, info in sorted(parties.items(), key=lambda kv: (kv[1]["name"] or kv[1]["inn"]).lower()):
            self.tree.insert("", "end", iid=key, values=("☐", info["inn"], info["name"], info["sample"]))

        footer = ttk.Frame(self, padding=14)
        footer.pack(fill="x")
        ttk.Label(footer, text="Guruh nomi:").pack(side="left")
        self.cat_var = tk.StringVar()
        ttk.Entry(footer, textvariable=self.cat_var, width=22).pack(side="left", padx=(6, 14))
        self.count_label = ttk.Label(footer, text="0 ta belgilandi")
        self.count_label.pack(side="left")
        ttk.Button(footer, text="Guruhga qo'shish", style="Accent.TButton", command=self._apply).pack(side="right")
        ttk.Button(footer, text="Yopish", command=self.destroy).pack(side="right", padx=(0, 8))

    def _on_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        row = self.tree.identify_row(event.y)
        if not row or self.tree.identify_column(event.x) != "#1":
            return
        if row in self.checked:
            self.checked.remove(row)
            self.tree.set(row, "check", "☐")
        else:
            self.checked.add(row)
            self.tree.set(row, "check", "☑")
        self.count_label.config(text=f"{len(self.checked)} ta belgilandi")

    def _apply(self):
        if not self.checked:
            messagebox.showwarning("Diqqat", "Kamida bitta qatorni belgilang.")
            return
        cat = self.cat_var.get().strip()
        if not cat:
            messagebox.showwarning("Diqqat", "Guruh nomini kiriting.")
            return
        for key in self.checked:
            info = self.parties[key]
            if info["inn"]:
                categorize.save_learned_category(info["inn"], "", cat)
            else:
                categorize.save_learned_category("", info["name"], cat)
        self.applied = True
        messagebox.showinfo("Saqlandi", f"{len(self.checked)} ta kontragent \"{cat}\" guruhiga qo'shildi.")
        self.destroy()


class App(tk.Tk):
    STATUS_COLORS = {
        "Kutmoqda": "#666666",
        "Ishlanmoqda...": "#0a66c2",
        "Tayyor": "#1a7f37",
        "Xato": "#c62828",
    }

    def __init__(self):
        super().__init__()
        self.title("Bank hisobotini soddalashtirish")
        self.geometry("960x680")
        self.minsize(820, 560)

        sv_ttk.set_theme("light")
        self._setup_fonts()

        self.files = []  # list[FileRow]
        self.out_dir = tk.StringVar(value="")
        self.is_running = False
        self.ui_queue = queue.Queue()

        self._build_ui()
        self.after(100, self._poll_queue)

    # ---------------------------------------------------------- UI layout
    def _setup_fonts(self):
        base = tkfont.nametofont("TkDefaultFont")
        base.configure(family="Segoe UI", size=10)
        self.option_add("*Font", base)
        self.heading_font = tkfont.Font(family="Segoe UI Semibold", size=17)
        self.subtitle_font = tkfont.Font(family="Segoe UI", size=10)
        self.step_font = tkfont.Font(family="Segoe UI Semibold", size=11)
        self.mono_font = tkfont.Font(family="Consolas", size=9)

    def _step_frame(self, parent, title):
        """A labeled 'card' section used to break the workflow into clear,
        numbered steps so a non-technical user always knows what's next."""
        frame = ttk.LabelFrame(parent, text=title, padding=14)
        return frame

    def _build_ui(self):
        pad = 14
        root = ttk.Frame(self, padding=pad)
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root)
        header.pack(fill="x", pady=(0, 14))
        ttk.Label(header, text="Bank hisobotini soddalashtirish", font=self.heading_font).pack(anchor="w")
        ttk.Label(
            header,
            text="Xom Hamkorbank hisobotlarini tanlang — har biri alohida Excel faylga aylantiriladi.",
            font=self.subtitle_font,
            foreground="#666666",
        ).pack(anchor="w")

        step1 = self._step_frame(root, "1-qadam · Fayllarni tanlang")
        step1.pack(fill="x", pady=(0, 10))
        row1 = ttk.Frame(step1)
        row1.pack(fill="x")
        ttk.Button(row1, text="📂 Fayllarni tanlash...", command=self.pick_files).pack(side="left")
        ttk.Button(row1, text="Ro'yxatni tozalash", command=self.clear_files).pack(side="left", padx=(8, 0))
        ttk.Button(row1, text="🏷 Guruhlarni boshqarish...", command=self.open_groups_manager).pack(side="left", padx=(8, 0))
        self.count_label = ttk.Label(row1, text="0 ta fayl tanlangan", font=self.step_font)
        self.count_label.pack(side="left", padx=(16, 0))

        step2 = self._step_frame(root, "2-qadam · Natijalarni qayerga saqlash")
        step2.pack(fill="x", pady=(0, 10))
        out_frame = ttk.Frame(step2)
        out_frame.pack(fill="x")
        self.out_entry = ttk.Entry(out_frame, textvariable=self.out_dir)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(out_frame, text="💾 Papka tanlash...", command=self.pick_out_dir).pack(side="left")

        step3 = self._step_frame(root, "3-qadam · Boshlang va kuzating")
        step3.pack(fill="both", expand=True)

        list_frame = ttk.Frame(step3)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))

        columns = ("file", "status")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="extended", height=10)
        self.tree.heading("file", text="Fayl")
        self.tree.heading("status", text="Holati")
        self.tree.column("file", width=600, anchor="w")
        self.tree.column("status", width=180, anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        scroll.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)

        self.tree.tag_configure("Kutmoqda", foreground=self.STATUS_COLORS["Kutmoqda"])
        self.tree.tag_configure("Ishlanmoqda...", foreground=self.STATUS_COLORS["Ishlanmoqda..."])
        self.tree.tag_configure("Tayyor", foreground=self.STATUS_COLORS["Tayyor"])
        self.tree.tag_configure("Xato", foreground=self.STATUS_COLORS["Xato"])
        ttk.Label(
            step3, text="Tayyor bo'lgan faylni ochish uchun ustiga ikki marta bosing.",
            foreground="#888888",
        ).pack(anchor="w", pady=(0, 8))

        self.progress = ttk.Progressbar(step3, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=(0, 10))

        action_row = ttk.Frame(step3)
        action_row.pack(fill="x")
        self.start_btn = ttk.Button(action_row, text="▶  Boshlash", style="Accent.TButton", command=self.start_processing)
        self.start_btn.pack(side="left", ipadx=6)
        self.open_out_btn = ttk.Button(action_row, text="📁 Papkani ochish", command=self.open_out_dir)
        self.open_out_btn.pack(side="left", padx=(8, 0))
        self.open_excel_btn = ttk.Button(action_row, text="📊 Excelda ochish", command=self.open_selected_in_excel)
        self.open_excel_btn.pack(side="left", padx=(8, 0))
        self.status_label = ttk.Label(action_row, text="Tayyor", font=self.step_font)
        self.status_label.pack(side="right")

        log_frame = ttk.LabelFrame(root, text="Jurnal", padding=(10, 6))
        log_frame.pack(fill="both", pady=(10, 0))
        self.log = tk.Text(
            log_frame, height=7, wrap="word", state="disabled",
            font=self.mono_font, relief="flat", borderwidth=0,
            background="#f5f5f5" if sv_ttk.get_theme() == "light" else "#1e1e1e",
        )
        self.log.pack(fill="both", expand=True)

    # ------------------------------------------------------------ actions
    def pick_files(self):
        paths = filedialog.askopenfilenames(
            title="Xom hisobot fayllarini tanlang",
            filetypes=[("Excel fayllar", "*.xlsx"), ("Barcha fayllar", "*.*")],
        )
        if not paths:
            return
        existing = {f.path for f in self.files}
        for p in paths:
            if p not in existing:
                self.files.append(FileRow(p))
        self._refresh_tree()

    def clear_files(self):
        if self.is_running:
            return
        self.files = []
        self._refresh_tree()

    def open_groups_manager(self):
        dialog = GroupsManagerDialog(self)
        self.wait_window(dialog)

    def pick_out_dir(self):
        d = filedialog.askdirectory(title="Natijalarni saqlash papkasini tanlang")
        if d:
            self.out_dir.set(d)

    def open_out_dir(self):
        d = self.out_dir.get().strip()
        if d and os.path.isdir(d):
            os.startfile(d)
        else:
            messagebox.showinfo("Diqqat", "Avval saqlash papkasini tanlang.")

    def _on_tree_double_click(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        f = self.files[idx]
        if f.out_path and os.path.exists(f.out_path):
            self._open_path_in_excel(f.out_path)

    def open_selected_in_excel(self):
        sel = self.tree.selection()
        if sel:
            targets = [self.files[int(i)] for i in sel]
        else:
            targets = self.files
        paths = [f.out_path for f in targets if f.out_path and os.path.exists(f.out_path)]
        if not paths:
            messagebox.showinfo(
                "Diqqat",
                "Ochish uchun tayyor natija fayli topilmadi. Avval \"Boshlash\" bilan qayta ishlashni yakunlang.",
            )
            return
        excel_exe = find_excel_exe()
        if not excel_exe:
            messagebox.showwarning(
                "Excel topilmadi",
                "Bu kompyuterda Microsoft Excel ro'yxatdan o'tmagan.\n\n"
                "Fayllar to'g'ri .xlsx formatida saqlangan, lekin ularni ochish uchun "
                "Excel (yoki LibreOffice Calc, WPS Office kabi mos dastur) o'rnatilgan bo'lishi kerak.\n\n"
                "Agar Excel o'rnatilgan bo'lsa-yu, fayl baribir Notepad'da ochilsa: fayl ustida "
                "o'ng tugmani bosing -> \"Open with\" -> Excel -> \"Always use this app\".",
            )
            return
        for p in paths:
            try:
                subprocess.Popen([excel_exe, p])
            except Exception as e:
                self._log(f"XATO: Excelda ochib bo'lmadi ({os.path.basename(p)}): {e}")

    def _open_path_in_excel(self, path):
        excel_exe = find_excel_exe()
        if excel_exe:
            try:
                subprocess.Popen([excel_exe, path])
                return
            except Exception as e:
                self._log(f"XATO: Excelda ochib bo'lmadi ({os.path.basename(path)}): {e}")
        os.startfile(path)

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i, f in enumerate(self.files):
            self.tree.insert("", "end", iid=str(i), values=(f.path, f.status), tags=(f.status,))
        self.count_label.config(text=f"{len(self.files)} ta fayl tanlangan")

    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def start_processing(self):
        if self.is_running:
            return
        if not self.files:
            messagebox.showwarning("Diqqat", "Avval kamida bitta fayl tanlang.")
            return
        out_dir = self.out_dir.get().strip()
        if not out_dir:
            messagebox.showwarning("Diqqat", "Avval saqlash papkasini tanlang.")
            return
        if not os.path.isdir(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Xato", f"Papka yaratib bo'lmadi:\n{e}")
                return

        for f in self.files:
            f.status = "Kutmoqda"
            f.error = None
        self._refresh_tree()

        self.is_running = True
        self.start_btn.configure(state="disabled")
        self.status_label.configure(text="Tekshirilmoqda...")
        self._log(f"--- Skanerlash boshlandi: {len(self.files)} ta fayl ---")

        t = threading.Thread(target=self._scan_worker, args=(out_dir,), daemon=True)
        t.start()

    def _scan_worker(self, out_dir):
        """Ishlov berishdan oldin: barcha tanlangan fayllarni o'qib, hech
        qanday kategoriyaga to'g'ri kelmagan ("?") kontragentlarni yig'ib
        chiqadi, shu bilan foydalanuvchidan bir marta so'rab olib bo'lgach
        haqiqiy qayta ishlash boshlanadi."""
        unresolved = {}
        for f in self.files:
            try:
                _wb, _ws, _hdr, rows = categorize.load_raw_rows(f.path)
            except Exception as e:
                self.ui_queue.put(("log", None, f"Skanerlashda xato ({os.path.basename(f.path)}): {e}", None))
                continue
            for key, info in categorize.find_unresolved(rows).items():
                if key not in unresolved:
                    unresolved[key] = dict(info)
                else:
                    unresolved[key]["count"] += info["count"]
        self.ui_queue.put(("scan_done", out_dir, unresolved, None))

    def _handle_scan_done(self, out_dir, unresolved):
        if unresolved:
            self._log(f"{len(unresolved)} ta nomlanmagan/kategoriyalanmagan kontragent topildi.")
            dialog = UnresolvedDialog(self, unresolved)
            self.wait_window(dialog)
            if not dialog.confirmed:
                self._log("Bekor qilindi.")
                self.is_running = False
                self.start_btn.configure(state="normal")
                self.status_label.configure(text="Tayyor")
                return
            assignments = dialog.get_assignments()
            for inn, name, cat in assignments:
                categorize.save_learned_category(inn, name, cat)
                self._log(f"Saqlandi: {name or inn} -> {cat}")
        else:
            self._log("Nomlanmagan kontragent topilmadi.")

        self.progress.configure(maximum=len(self.files), value=0)
        self.status_label.configure(text="Ishlanmoqda...")
        self._log(f"--- Qayta ishlash boshlandi: {len(self.files)} ta fayl ---")
        t = threading.Thread(target=self._worker, args=(out_dir,), daemon=True)
        t.start()

    def _worker(self, out_dir):
        done_ok = 0
        done_err = 0
        for idx, f in enumerate(self.files):
            self.ui_queue.put(("status", idx, "Ishlanmoqda...", None))
            base = os.path.splitext(os.path.basename(f.path))[0]
            out_path = os.path.join(out_dir, f"{base} - soddalashtirilgan.xlsx")
            try:
                info = build_report.run(f.path, out_path)
                f.out_path = out_path
                self.ui_queue.put(("status", idx, "Tayyor", None))
                self.ui_queue.put((
                    "log", None,
                    f"OK: {os.path.basename(f.path)} -> {os.path.basename(out_path)} "
                    f"({info['review_count']} ta tekshirish bandi)",
                    None,
                ))
                done_ok += 1
            except Exception as e:
                err = "".join(traceback.format_exception_only(type(e), e)).strip()
                self.ui_queue.put(("status", idx, "Xato", err))
                self.ui_queue.put(("log", None, f"XATO: {os.path.basename(f.path)} -> {err}", None))
                done_err += 1
            self.ui_queue.put(("progress", idx + 1, None, None))
        self.ui_queue.put(("done", done_ok, done_err, None))

    def _poll_queue(self):
        try:
            while True:
                kind, a, b, _c = self.ui_queue.get_nowait()
                if kind == "status":
                    idx, status = a, b
                    self.files[idx].status = status
                    self.tree.item(str(idx), values=(self.files[idx].path, status), tags=(status,))
                elif kind == "progress":
                    self.progress.configure(value=a)
                elif kind == "log":
                    self._log(b)
                elif kind == "scan_done":
                    out_dir, unresolved = a, b
                    self._handle_scan_done(out_dir, unresolved)
                elif kind == "done":
                    ok, err = a, b
                    self.is_running = False
                    self.start_btn.configure(state="normal")
                    self.status_label.configure(text="Tayyor")
                    self._log(f"--- Jarayon tugadi: {ok} ta muvaffaqiyatli, {err} ta xato ---")
                    tip = (
                        "\n\nFaylni ochish uchun ustiga ikki marta bosing yoki "
                        "\"Excelda ochish\" tugmasini bosing (agar oddiy ikki marta "
                        "bosish Notepad'da ochsa, bu tugma majburan Excel bilan ochadi)."
                        if ok else ""
                    )
                    if err:
                        messagebox.showwarning(
                            "Tugadi",
                            f"{ok} ta fayl tayyor bo'ldi, {err} ta faylda xato yuz berdi.\nJurnalni tekshiring.{tip}",
                        )
                    else:
                        messagebox.showinfo("Tugadi", f"Barcha {ok} ta fayl muvaffaqiyatli qayta ishlandi.{tip}")
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)


def main():
    os.chdir(APP_DIR)
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

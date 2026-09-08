# Tkinter UI qo'llanmasi — boshqa loyihalarda qayta ishlatish uchun

Bu hujjat SoddaHisobot ilovasida qo'llangan barcha UI yechimlarini to'playdi.
Har bir bo'limda: **muammo → sabab → yechim (kod)**.

Stack: `tkinter` + `ttk` + [`sv-ttk`](https://pypi.org/project/sv-ttk/) (Windows 11 uslubidagi tema).

```
pip install sv-ttk
```

PyInstaller bilan yig'ilganda `sv_ttk` fayllari avtomatik qo'shiladi —
`pyinstaller-hooks-contrib` da tayyor hook bor, `--collect-data` shart emas.

---

## 1. Poydevor: DPI, tema, shrift

### Muammo
Windows monitor masshtabi 125%/150% bo'lganda yoki oyna boshqa monitorga
ko'chirilganda elementlar siljib, ustma-ust tushib qolardi.

### Yechim
Dasturni ishga tushirishdan **oldin** Windows'ga "men DPI-ga moslashaman"
deb aytish kerak, keyin Tk masshtabini real DPI ga sozlash.

```python
def _enable_dpi_awareness():
    """main() da, Tk() yaratilishidan OLDIN chaqiriladi."""
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)   # PROCESS_SYSTEM_DPI_AWARE
    except Exception:
        try:
            from ctypes import windll
            windll.user32.SetProcessDPIAware()     # eski Windows uchun
        except Exception:
            pass

def main():
    _enable_dpi_awareness()   # <-- Tk() dan oldin
    App().mainloop()
```

Tk ichida esa real DPI ga qarab masshtab:

```python
def _apply_dpi_scaling(self):
    try:
        dpi = self.winfo_fpixels("1i")
        if dpi > 0:
            self.tk.call("tk", "scaling", dpi / 72.0)
    except Exception:
        pass
```

### Tema va shriftlar

```python
import sv_ttk
from tkinter import font as tkfont

sv_ttk.set_theme("light")        # yoki "dark"

def _setup_fonts(self):
    base = tkfont.nametofont("TkDefaultFont")
    base.configure(family="Segoe UI", size=10)
    self.option_add("*Font", base)          # butun ilovaga tarqaladi
    self.heading_font  = tkfont.Font(family="Segoe UI Semibold", size=17)
    self.subtitle_font = tkfont.Font(family="Segoe UI", size=10)
    self.step_font     = tkfont.Font(family="Segoe UI Semibold", size=11)
    self.mono_font     = tkfont.Font(family="Consolas", size=9)
```

sv-ttk **`Accent.TButton`** uslubini beradi — asosiy harakat tugmasi uchun:

```python
ttk.Button(row, text="Boshlash", style="Accent.TButton", command=...)
```

---

## 2. Bosqichli ("qadam") joylashuv

### Muammo
Texnik bo'lmagan foydalanuvchi nimadan boshlashni bilmaydi.

### Yechim
Har bir bosqich alohida `ttk.LabelFrame` — sarlavhasi raqamlangan "karta".

```python
def _step_frame(self, parent, title):
    return ttk.LabelFrame(parent, text=title, padding=14)

step1 = self._step_frame(root, "1-qadam - Fayllarni tanlang")
step1.pack(fill="x", pady=(0, 10))

step2 = self._step_frame(root, "2-qadam - Natijalarni qayerga saqlash")
step2.pack(fill="x", pady=(0, 10))

step3 = self._step_frame(root, "3-qadam - Boshlang va kuzating")
step3.pack(fill="both", expand=True)     # faqat shu kengayadi
```

Tugmalarga emoji qo'shish tushunarlilikni oshiradi, kod esa o'zgarmaydi.

---

## 3. ENG MUHIM: tugmalar kesilib qolmasligi

### Muammo
Oyna kichraytirilganda pastdagi tugmalar ko'rinmay qolardi — faqat to'liq
ekranga yoyilganda ko'rinardi.

### Sabab
`pack` bo'sh joyni **paketlash tartibida** taqsimlaydi. Agar kengayuvchi
ro'yxat (`expand=True`) birinchi paketlansa, u butun joyni oladi va keyin
paketlangan tugmalarga joy qolmaydi.

### Yechim
Doimiy balandlikdagi elementlarni **birinchi** va `side="bottom"` bilan
paketlash kerak. Kod tartibini o'zgartirmaslik uchun `before=` ishlatiladi:

```python
# NOTO'G'RI - tugmalar kesiladi
tree_frame.pack(fill="both", expand=True)
footer.pack(fill="x")

# TO'G'RI - 1-variant: footer'ni oldin paketlash
footer.pack(side="bottom", fill="x")
tree_frame.pack(fill="both", expand=True)

# TO'G'RI - 2-variant: kod tartibi o'zgarmaydi, before= yordam beradi
tree_frame.pack(fill="both", expand=True)
footer.pack(side="bottom", fill="x", before=tree_frame)
```

Qo'shimcha: `Treeview`/`Text` ning `height=` sini kichik qiling (masalan
`height=4`) — bu **minimal** balandlikni belgilaydi, `expand=True` esa joy
bo'lsa baribir cho'zadi.

### Tekshirish usuli
Ko'z bilan emas, o'lchab tekshiring:

```python
a = App(); a.geometry("780x520"); a.update_idletasks(); a.update()
pastki = widget.winfo_rooty() - a.winfo_rooty() + widget.winfo_height()
print("KO'RINADI" if pastki <= a.winfo_height() else "KESILGAN")
```

---

## 4. Maximize/restore da qora "yamoq"lar

### Muammo
Oyna to'liq ekranga yoyilganda ba'zi joylar qora bo'lib qolardi.

### Sabab
Tk/DWM darajasidagi qayta chizish nuqsoni — rasm asosidagi temalarda
(sv-ttk) ko'proq uchraydi.

### Yechim
Oyna **holati** o'zgarganda (normal <-> zoomed) butun daraxtni majburan
qayta chizish. Har bir piksel o'zgarishida emas — faqat holat almashganda:

```python
self._last_state = self.state()
self.bind("<Configure>", self._on_root_configure)

def _on_root_configure(self, event):
    if event.widget is not self:
        return
    try:
        state = self.state()
    except tk.TclError:
        return
    if state != self._last_state:
        self._last_state = state
        self.after(50, self._force_full_redraw)

def _force_full_redraw(self):
    def redraw(w):
        w.update_idletasks()
        for child in w.winfo_children():
            redraw(child)
    try:
        redraw(self)
    except tk.TclError:
        pass
```

---

## 5. Matn oyna kengligiga moslashishi

`ttk.Label` uzun matni avtomatik o'ralmaydi — `wraplength` ni oyna
kengligiga bog'lash kerak:

```python
root.bind("<Configure>", self._on_root_resize)

def _on_root_resize(self, event):
    try:
        self.subtitle_label.configure(wraplength=max(300, event.width - 28))
    except Exception:
        pass
```

---

## 6. Excel uslubidagi jadval — checkbox bilan

### Muammo
`ttk.Treeview` da haqiqiy checkbox yo'q.

### Yechim
Birinchi ustunga belgi yozib, faqat o'sha ustunga bosilganda almashtirish.
`selectmode="none"` — tanlash emas, belgilash kerak.

```python
columns = ("check", "account", "mfo", "name", "sample")
self.tree = ttk.Treeview(frame, columns=columns, show="headings",
                         selectmode="none", height=6)
self.tree.heading("check", text="v")
self.tree.column("check", width=36, anchor="center")
self.tree.bind("<Button-1>", self._on_click)

for key, info in data.items():
    self.tree.insert("", "end", iid=key, values=("[ ]", ...))

def _on_click(self, event):
    if self.tree.identify_region(event.x, event.y) != "cell":
        return
    row = self.tree.identify_row(event.y)
    if not row or self.tree.identify_column(event.x) != "#1":   # 1-ustun
        return
    if row in self.checked:
        self.checked.remove(row); self.tree.set(row, "check", "[ ]")
    else:
        self.checked.add(row);    self.tree.set(row, "check", "[x]")
```

### Holat ranglari (tag)

```python
self.tree.tag_configure("Tayyor", foreground="#1a7f37")
self.tree.tag_configure("Xato",   foreground="#c62828")
self.tree.insert("", "end", iid="1", values=(...), tags=("Tayyor",))
```

---

## 7. Aylantiriladigan (scrollable) dialog

`Frame` o'zi aylanmaydi — `Canvas` ichiga joylash kerak:

```python
canvas = tk.Canvas(self, highlightthickness=0)
scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
inner = ttk.Frame(canvas)
inner.bind("<Configure>",
           lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.create_window((0, 0), window=inner, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="left", fill="y")

# endi widget'lar `inner` ichiga qo'yiladi
```

---

## 8. Modal dialog patterni

```python
class MyDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Sarlavha")
        self.geometry("720x560")
        self.minsize(600, 420)       # kichraytirish chegarasi
        self.transient(parent)        # ota oyna ustida turadi
        self.grab_set()               # modal - ortidagi oyna bloklanadi
        self.result = None
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)  # X tugmasi

# Chaqirish:
dlg = MyDialog(self)
self.wait_window(dlg)      # yopilishini kutadi
if dlg.result:
    ...
```

---

## 9. UI qotib qolmasligi: thread + queue

### Muammo
Og'ir ish (fayl qayta ishlash, tarmoq) UI oqimida bajarilsa, oyna
"Not Responding" bo'lib qoladi.

### Yechim
Ish alohida `Thread` da, natija `queue.Queue` orqali UI ga uzatiladi, UI esa
`after()` bilan navbatni tekshirib turadi. **Widget'ga faqat UI oqimidan
tegiladi.**

```python
import queue, threading

self.ui_queue = queue.Queue()
self.after(100, self._poll_queue)

def _boshlash(self):
    threading.Thread(target=self._worker, daemon=True).start()

def _worker(self):                       # <-- boshqa oqim
    for i, item in enumerate(items):
        natija = ogir_ish(item)
        self.ui_queue.put(("progress", i + 1, None))
        self.ui_queue.put(("log", f"{item} tayyor", None))
    self.ui_queue.put(("done", None, None))

def _poll_queue(self):                   # <-- UI oqimi
    try:
        while True:
            kind, a, b = self.ui_queue.get_nowait()
            if kind == "progress":
                self.progress.configure(value=a)
            elif kind == "log":
                self._log(a)
            elif kind == "done":
                messagebox.showinfo("Tayyor", "Yakunlandi")
    except queue.Empty:
        pass
    self.after(100, self._poll_queue)
```

Fon oqimidan xabar berish (masalan telemetriya) — hech narsani kutmasdan:

```python
threading.Thread(target=send_ping, daemon=True).start()
```

---

## 10. Jurnal (log) maydoni

```python
self.log = tk.Text(frame, height=4, wrap="word", state="disabled",
                   font=self.mono_font, relief="flat", borderwidth=0,
                   background="#f5f5f5")

def _log(self, msg):
    self.log.configure(state="normal")     # yozish uchun ochamiz
    self.log.insert("end", msg + "\n")
    self.log.see("end")                    # oxiriga aylantiramiz
    self.log.configure(state="disabled")   # foydalanuvchi o'zgartira olmaydi
```

Foydali: ilova ochilganda birinchi qatorga versiya va `.exe` yo'lini yozing —
"menda eski versiya ko'rinyapti" muammosini bir zumda hal qiladi:

```python
self._log(f"Versiya v{APP_VERSION}  |  Joylashuvi: {os.path.abspath(sys.executable)}")
```

---

## 11. Windows bilan integratsiya

### Faylni majburan Excel bilan ochish
`.xlsx` Notepad'ga bog'lanib qolgan bo'lsa ham ishlaydi — registrdan haqiqiy
Excel yo'li olinadi:

```python
import winreg, subprocess, os

def find_excel_exe():
    candidates = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE"),
        (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE"),
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

subprocess.Popen([find_excel_exe(), fayl_yoli])
```

### Ish stoliga yorliq (PowerShell orqali, qo'shimcha kutubxonasiz)

```python
ps = (
    "$W = New-Object -ComObject WScript.Shell; "
    f'$S = $W.CreateShortcut("{shortcut_path}"); '
    f'$S.TargetPath = "{target}"; $S.WorkingDirectory = "{workdir}"; '
    f'$S.IconLocation = "{target}"; $S.Save()'
)
subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
               creationflags=subprocess.CREATE_NO_WINDOW, timeout=10, check=False)
```

---

## Yakuniy tekshiruv ro'yxati

Yangi ilova qilganda shularni tekshiring:

- [ ] `_enable_dpi_awareness()` — `Tk()` dan **oldin** chaqirilganmi
- [ ] `sv_ttk.set_theme(...)` va `_setup_fonts()` qo'yilganmi
- [ ] Doimiy balandlikdagi qismlar `side="bottom"` bilan **oldin**
      paketlanganmi (`before=` yoki kod tartibi bilan)
- [ ] `Treeview`/`Text` ning `height=` si kichikmi (minimal o'lcham)
- [ ] `minsize()` da barcha tugmalar ko'rinadimi — **o'lchab** tekshiring
- [ ] Maximize/restore da qora joylar chiqmaydimi
- [ ] Og'ir ish alohida `Thread` da, UI ga faqat `queue` orqali tegilyaptimi
- [ ] Dialoglar `transient()` + `grab_set()` + `WM_DELETE_WINDOW` bilan

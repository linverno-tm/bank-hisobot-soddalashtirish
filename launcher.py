# -*- coding: utf-8 -*-
"""
SoddaHisobot — LAUNCHER (qobiq).

Bu fayl .exe ga kompilyatsiya qilinadi va foydalanuvchida shu turadi.
Ilovaning haqiqiy mantiqi bu yerda EMAS — u GitHub'dagi core.py da.
Launcher har ishga tushganda core.py ning eng yangi nusxasini yuklab
olib bajaradi.

Nima uchun shunday: tuzatish yoki yangi funksiya chiqarish uchun
core.py ni GitHub'ga push qilish kifoya. Foydalanuvchi ilovani keyingi
ochganda yangi kodni oladi — .exe ni qayta yuklab olish, o'rnatish,
antivirus ogohlantirishlari, "eski nusxalar" muammosi — hech biri yo'q.

DIQQAT: bu fayl deyarli hech qachon o'zgarmasligi kerak. U o'zini
yangilamaydi — o'zgartirilsa, .exe ni qayta yig'ib qo'lda tarqatish
kerak bo'ladi. Shuning uchun bu yerda hech qanday bog'liqlik yo'q,
faqat standart kutubxona: yuklab ol -> tekshir -> bajar.

Uch pog'onali zaxira: internet -> kesh -> .exe ichidagi nusxa.
Shuning uchun ilova internet bo'lmasa ham ochiladi.
"""
import os
import sys
import types
import urllib.request

APP_NAME = "SoddaHisobot"
CORE_URL = "https://raw.githubusercontent.com/linverno-tm/bank-hisobot-soddalashtirish/{branch}/core.py"
CORE_FILE = "core.py"
BRANCH_FILE = "branch.txt"
DEFAULT_BRANCH = "master"
DOWNLOAD_TIMEOUT = 6  # soniya — sekin internetda ilova ochilishini kutib qolmasin

# branch.txt ichidagi matn URL ga qo'shiladi, shuning uchun faqat
# shoxobcha nomida uchraydigan belgilarga ruxsat beriladi.
BRANCH_CHARS = set(
    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_./"
)

# Yuklangan matn haqiqiy core.py ekanini tasdiqlovchi belgilar. Yarim
# yuklangan yoki GitHub'ning xato sahifasi (HTML) bajarilib ketmasin.
SANITY_MARKERS = ("def main(", "CORE_VERSION")
MIN_SIZE = 5000  # bayt


def _cache_dir():
    """Yuklangan core.py saqlanadigan papka. %LOCALAPPDATA% ishlatiladi —
    .exe yonidagi papka "Program Files" da bo'lsa yozish taqiqlangan
    bo'lishi mumkin."""
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    path = os.path.join(base, APP_NAME)
    try:
        os.makedirs(path, exist_ok=True)
    except OSError:
        return None
    return path


def _branch():
    """Sinov rejimi. Kesh papkasida branch.txt bo'lsa va ichida shoxobcha
    nomi yozilgan bo'lsa, kod master'dan emas, o'sha shoxobchadan olinadi.

    Shu tufayli bitta .exe ham sinov, ham ishchi versiya bo'la oladi:
    dasturchining kompyuterida branch.txt bor, foydalanuvchida yo'q —
    tugallanmagan kod push qilinsa ham unga bormaydi."""
    d = _cache_dir()
    if not d:
        return DEFAULT_BRANCH
    nom = (_from_file(os.path.join(d, BRANCH_FILE)) or "").strip()
    if not nom or ".." in nom or nom.startswith((".", "/")):
        return DEFAULT_BRANCH
    if not set(nom) <= BRANCH_CHARS:
        return DEFAULT_BRANCH
    return nom


def _cache_path():
    d = _cache_dir()
    if not d:
        return None
    shox = _branch()
    if shox == DEFAULT_BRANCH:
        return os.path.join(d, CORE_FILE)
    # Har shoxobchaning keshi alohida saqlanadi — sinovdan qaytilganda
    # ishchi versiyaning keshi buzilmagan holda joyida qoladi.
    return os.path.join(d, "core-%s.py" % shox.replace("/", "-"))


def _bundled_path():
    """.exe ichiga --add-data bilan o'ralgan zaxira nusxa."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, CORE_FILE)


def _is_valid(source):
    return (
        source
        and len(source) >= MIN_SIZE
        and all(m in source for m in SANITY_MARKERS)
    )


def _from_url():
    req = urllib.request.Request(
        CORE_URL.format(branch=_branch()),
        headers={"User-Agent": f"{APP_NAME}-Launcher", "Cache-Control": "no-cache"},
    )
    with urllib.request.urlopen(req, timeout=DOWNLOAD_TIMEOUT) as resp:
        return resp.read().decode("utf-8")


def _from_file(path):
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as f:
            return f.read()
    except OSError:
        return None


def _save_cache(source):
    path = _cache_path()
    if not path:
        return
    try:
        # Avval vaqtinchalik faylga yozib, keyin o'rniga qo'yamiz —
        # yozish yarmida to'xtasa, keshda buzuq fayl qolib ketmaydi.
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            f.write(source)
        os.replace(tmp, path)
    except OSError:
        pass


def _load(source, shoxobcha):
    """Matnni modulga aylantiradi va main() ni qaytaradi — lekin HALI
    chaqirmaydi. Shu bosqichda sintaksis xatolari, yo'q kutubxonalar va
    modul darajasidagi xatolar aniqlanadi. Ya'ni "kod ishga yaroqlimi"
    degan savolga javob shu yerda olinadi."""
    module = types.ModuleType("core")
    module.__dict__["__file__"] = _cache_path() or _bundled_path()
    # core.py shu qiymatga qarab sarlavhada sinov rejimini ko'rsatadi —
    # sinov nusxasini ishchi nusxa deb o'ylab qolmaslik uchun.
    module.__dict__["SOURCE_BRANCH"] = shoxobcha
    exec(compile(source, "core.py", "exec"), module.__dict__)
    main = module.__dict__.get("main")
    if not callable(main):
        raise RuntimeError("core.py ichida main() topilmadi")
    sys.modules["core"] = module
    return main


def _run(source, shoxobcha):
    """Kodni yuklab, ishga tushiradi."""
    _load(source, shoxobcha)()


def _show_error(matn):
    """Grafik ilova bo'lgani uchun konsol yo'q — xatoni oyna orqali
    ko'rsatamiz, aks holda .exe jimgina yopilib qoladi."""
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_NAME, matn)
        root.destroy()
    except Exception:
        pass


def _pyinstaller_bogliqliklari():
    """HECH QACHON CHAQIRILMAYDI — faqat PyInstaller uchun.

    core.py exec() ichida bajariladi, shuning uchun PyInstaller uning
    import'larini ko'ra olmaydi va ularni .exe ga qo'shmaydi. Natijada
    ilova foydalanuvchida "No module named ..." bilan yiqiladi. Bu
    ro'yxat esa oddiy import bo'lgani uchun PyInstaller uni ko'radi.

    Diqqat: standart kutubxona ham kerak (json, re, decimal...) — .exe ga
    butun Python emas, faqat ko'ringan modullar tushadi.

    core.py ga YANGI import qo'shilsa, shu yerga ham qo'shib, .exe ni
    qayta yig'ish kerak. Aks holda yangi kod GitHub'dan tushadi-yu,
    eski .exe da ishlamaydi."""
    import collections  # noqa: F401
    import copy  # noqa: F401
    import datetime  # noqa: F401
    import decimal  # noqa: F401
    import json  # noqa: F401
    import platform  # noqa: F401
    import queue  # noqa: F401
    import re  # noqa: F401
    import subprocess  # noqa: F401
    import threading  # noqa: F401
    import traceback  # noqa: F401
    import winreg  # noqa: F401

    import openpyxl  # noqa: F401
    import openpyxl.cell.cell  # noqa: F401
    import openpyxl.styles  # noqa: F401
    import openpyxl.utils  # noqa: F401
    import sv_ttk  # noqa: F401

    import tkinter  # noqa: F401
    import tkinter.filedialog  # noqa: F401
    import tkinter.font  # noqa: F401
    import tkinter.messagebox  # noqa: F401
    import tkinter.ttk  # noqa: F401


def main():
    shox = _branch()

    # 1) Internetdan eng yangi nusxa
    yangi = None
    try:
        yangi = _from_url()
    except Exception:
        yangi = None

    if _is_valid(yangi):
        try:
            core_main = _load(yangi, shox)
        except Exception:
            # Yangi kod buzuq (push'da xato): keshga TEGMAYMIZ, aks holda
            # oxirgi ishlagan nusxa ham yo'qolib, foydalanuvchi eski
            # zaxiraga tushib qolardi.
            core_main = None
        if core_main is not None:
            # Kod yuklanishga yarokli ekani tasdiqlandi — endi keshni
            # yangilaymiz. Shundan keyingi xatolar (interfeys ichida)
            # keshga ta'sir qilmaydi.
            _save_cache(yangi)
            core_main()
            return

    # 2) Keshdagi oxirgi ishlagan nusxa
    keshdagi = _from_file(_cache_path())
    if _is_valid(keshdagi) and keshdagi != yangi:
        try:
            _run(keshdagi, shox)
            return
        except Exception:
            pass

    # 3) .exe ichidagi zaxira nusxa. U .exe yig'ilgan paytdagi master
    # nusxasi — sinov shoxobchasiga aloqasi yo'q.
    zaxira = _from_file(_bundled_path())
    if _is_valid(zaxira):
        try:
            _run(zaxira, DEFAULT_BRANCH)
            return
        except Exception as e:
            _show_error(f"Ilovani ishga tushirib bo'lmadi:\n\n{e}")
            return

    _show_error(
        "Ilova kodini yuklab bo'lmadi.\n\n"
        "Internetga ulanishni tekshirib, ilovani qayta ochib ko'ring."
    )


if __name__ == "__main__":
    main()

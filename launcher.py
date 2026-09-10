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
CORE_URL = "https://raw.githubusercontent.com/linverno-tm/bank-hisobot-soddalashtirish/master/core.py"
CORE_FILE = "core.py"
DOWNLOAD_TIMEOUT = 6  # soniya — sekin internetda ilova ochilishini kutib qolmasin

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


def _cache_path():
    d = _cache_dir()
    return os.path.join(d, CORE_FILE) if d else None


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
        CORE_URL,
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


def _run(source, manba):
    """Matnni modulga aylantirib, main() ni chaqiradi."""
    module = types.ModuleType("core")
    module.__dict__["__file__"] = _cache_path() or _bundled_path()
    exec(compile(source, "core.py", "exec"), module.__dict__)
    main = module.__dict__.get("main")
    if not callable(main):
        raise RuntimeError("core.py ichida main() topilmadi")
    sys.modules["core"] = module
    main()


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


def main():
    # 1) Internetdan eng yangi nusxa
    yangi = None
    try:
        yangi = _from_url()
    except Exception:
        yangi = None

    if _is_valid(yangi):
        _save_cache(yangi)
        try:
            _run(yangi, "internet")
            return
        except Exception:
            # Yangi kod buzuq bo'lsa (push'da xato) — ilova butunlay
            # ishlamay qolmasligi uchun ishlagan nusxaga qaytamiz.
            pass

    # 2) Keshdagi oxirgi ishlagan nusxa
    keshdagi = _from_file(_cache_path())
    if _is_valid(keshdagi) and keshdagi != yangi:
        try:
            _run(keshdagi, "kesh")
            return
        except Exception:
            pass

    # 3) .exe ichidagi zaxira nusxa
    zaxira = _from_file(_bundled_path())
    if _is_valid(zaxira):
        try:
            _run(zaxira, "zaxira")
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

# -*- coding: utf-8 -*-
"""
Ilovaning o'zini GitHub Releases'dan tekshirib, yangi versiya chiqqan bo'lsa
avtomatik yuklab, ishlab turgan .exe faylni almashtiradigan modul.

Ishlash tartibi:
  1. check_for_update() — GitHub API orqali eng oxirgi release'ni so'raydi,
     versiyani taqqoslaydi. Internet yo'q yoki GitHub javob bermasa, jim
     tarzda None qaytaradi (ilova ishlashiga hech qanday xalaqit bermaydi).
  2. download_and_apply_update() — yangi .exe (yoki uni ichida saqlagan
     .zip) ni yuklab oladi, joriy .exe'ni "_old_" prefiksi bilan qayta
     nomlab, yangisini uning o'rniga qo'yadi va to'g'ridan-to'g'ri ishga
     tushiradi. Hech qanday .bat/cmd.exe ishlatilmaydi — antiviruslar
     shunday sxemani zararli deb bloklaydi. persist_dictionary.json faylga
     HECH QACHON tegilmaydi — faqat .exe almashtiriladi, shuning uchun
     o'rgatilgan guruhlar yo'qolmaydi.
  3. cleanup_old_versions() — keyingi ishga tushishda qolib ketgan
     "_old_*.exe" nusxasini o'chiradi.

Bu modul faqat PyInstaller bilan yig'ilgan (frozen) .exe holatida ishlaydi;
oddiy "python app.py" orqali ishga tushirilganda yangilanish o'zini
o'chiradi (dasturchi versiyasini tasodifan almashtirib qo'ymaslik uchun).
"""
import json
import os
import subprocess
import sys
import zipfile
import urllib.request
import urllib.error

APP_VERSION = "1.9.0"
GITHUB_REPO = "linverno-tm/bank-hisobot-soddalashtirish"
_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
_USER_AGENT = "SoddaHisobot-Updater"
_OLD_EXE_PREFIX = "_old_"


def _version_tuple(v):
    out = []
    for part in str(v).strip().lstrip("vV").split("."):
        digits = "".join(ch for ch in part if ch.isdigit())
        out.append(int(digits) if digits else 0)
    return tuple(out)


def _is_newer(remote, local):
    r, l = _version_tuple(remote), _version_tuple(local)
    n = max(len(r), len(l))
    r = r + (0,) * (n - len(r))
    l = l + (0,) * (n - len(l))
    return r > l


def _pick_asset(assets):
    """Prefer a direct .exe release asset; fall back to a .zip that should
    contain the .exe inside it."""
    exe_asset = None
    zip_asset = None
    for a in assets:
        name = (a.get("name") or "").lower()
        url = a.get("browser_download_url")
        if not url:
            continue
        if name.endswith(".exe") and exe_asset is None:
            exe_asset = (url, a.get("name"))
        elif name.endswith(".zip") and zip_asset is None:
            zip_asset = (url, a.get("name"))
    return exe_asset or zip_asset


def check_for_update(timeout=5):
    """Tekshiradi va uchta holatdan birini qaytaradi (hech qachon xato
    ko'tarmaydi):
      - (tag_name, asset_url, asset_name, release_notes) — yangi versiya bor;
      - False — tekshiruv muvaffaqiyatli o'tdi va eng oxirgi versiya
        allaqachon o'rnatilgan;
      - None — tekshirib bo'lmadi (dev muhiti, internet yo'q, GitHub javob
        bermadi/limitga tushdi va h.k.) — bu holat "eng oxirgi versiya"
        bilan ADASHTIRILMASLIGI kerak, aks holda haqiqiy xatolik chog'ida
        foydalanuvchiga yolg'on "yangilanish yo'q" xabari ko'rsatiladi."""
    if not getattr(sys, "frozen", False):
        return None  # dev muhitida (python app.py) yangilanish tekshirilmaydi
    try:
        req = urllib.request.Request(
            _API_URL,
            headers={"Accept": "application/vnd.github+json", "User-Agent": _USER_AGENT},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None

    tag = str(data.get("tag_name") or "").strip()
    if not tag or not _is_newer(tag, APP_VERSION):
        return False

    asset = _pick_asset(data.get("assets") or [])
    if not asset:
        return None  # yangi tag bor-u, lekin yuklab bo'lmaydigan holat — xato sifatida ko'rsatamiz
    asset_url, asset_name = asset
    return tag, asset_url, asset_name, str(data.get("body") or "").strip()


def download_and_apply_update(asset_url, asset_name, progress_cb=None):
    """Yangi versiyani yuklab, joriy .exe o'rniga almashtirishni
    rejalashtiradi. Chaqiruvchi shundan keyin darhol dasturdan chiqishi
    kerak (masalan os._exit(0)), aks holda .bat skript eski faylni
    almashtira olmaydi (fayl band bo'lib qoladi)."""
    if not getattr(sys, "frozen", False):
        raise RuntimeError("Yangilash faqat build qilingan .exe versiyasida ishlaydi.")

    exe_path = os.path.abspath(sys.executable)
    exe_dir = os.path.dirname(exe_path)
    ext = os.path.splitext(asset_name)[1].lower() or ".bin"
    download_path = os.path.join(exe_dir, "_update_download" + ext)

    req = urllib.request.Request(asset_url, headers={"User-Agent": _USER_AGENT})
    with urllib.request.urlopen(req, timeout=60) as resp:
        total = int(resp.headers.get("Content-Length", 0) or 0)
        downloaded = 0
        with open(download_path, "wb") as f:
            while True:
                chunk = resp.read(65536)
                if not chunk:
                    break
                f.write(chunk)
                downloaded += len(chunk)
                if progress_cb:
                    progress_cb(downloaded, total)

    if download_path.lower().endswith(".zip"):
        new_exe_path = os.path.join(exe_dir, "_update_new.exe")
        with zipfile.ZipFile(download_path) as zf:
            exe_entry = next((n for n in zf.namelist() if n.lower().endswith(".exe")), None)
            if not exe_entry:
                raise RuntimeError("Yangilanish arxivida .exe fayl topilmadi.")
            with zf.open(exe_entry) as src, open(new_exe_path, "wb") as dst:
                dst.write(src.read())
        os.remove(download_path)
    else:
        new_exe_path = download_path

    # .bat + cmd.exe ishlatmaymiz: ishlab turgan .exe'ni almashtirish uchun
    # yordamchi skript ochish antiviruslar tomonidan zararli xatti-harakat
    # sifatida bloklanadi ("Security validation failure: failed to obtain
    # executable path for parent process" kabi xatolar shundan chiqadi).
    #
    # Windows ishlab turgan .exe faylni O'CHIRISHGA ruxsat bermaydi, lekin
    # QAYTA NOMLASHGA ruxsat beradi. Shundan foydalanamiz:
    #   1. joriy .exe -> "<nom>_old.exe" deb qayta nomlanadi
    #   2. yangi .exe uning o'rniga qo'yiladi
    #   3. yangi .exe to'g'ridan-to'g'ri ishga tushiriladi (hech qanday
    #      oraliq skriptsiz), joriy jarayon esa chiqib ketadi
    # Eski "_old.exe" keyingi ishga tushishda cleanup_old_versions() bilan
    # o'chiriladi (o'shanda u endi band bo'lmaydi).
    old_path = os.path.join(exe_dir, f"{_OLD_EXE_PREFIX}{os.path.basename(exe_path)}")
    if os.path.exists(old_path):
        try:
            os.remove(old_path)
        except OSError:
            pass

    os.replace(exe_path, old_path)
    try:
        os.replace(new_exe_path, exe_path)
    except OSError:
        os.replace(old_path, exe_path)  # muvaffaqiyatsiz bo'lsa, eskisini tiklaymiz
        raise

    subprocess.Popen([exe_path], cwd=exe_dir, close_fds=True)


def cleanup_old_versions():
    """Yangilanishdan keyin qolib ketgan eski .exe nusxasini o'chiradi.
    Ilova ishga tushganda chaqiriladi — o'shanda eski fayl endi band
    emas. Xato chiqsa jim o'tkazib yuboriladi (bu shunchaki tozalash)."""
    if not getattr(sys, "frozen", False):
        return
    try:
        exe_path = os.path.abspath(sys.executable)
        exe_dir = os.path.dirname(exe_path)
        old_path = os.path.join(exe_dir, f"{_OLD_EXE_PREFIX}{os.path.basename(exe_path)}")
        if os.path.exists(old_path):
            os.remove(old_path)
    except Exception:
        pass

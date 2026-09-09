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

APP_VERSION = "99.0.0"  # SINOV NUSXASI: master relizlariga yangilanib, AI yo'qolib ketmasligi uchun ataylab yuqori
GITHUB_REPO = "linverno-tm/bank-hisobot-soddalashtirish"
_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
_USER_AGENT = "SoddaHisobot-Updater"
_OLD_EXE_PREFIX = "_old_"
# Ilova ochilganda o'z versiyasini yuboradigan manzil. Faqat kompyuter
# nomi va versiya yuboriladi — boshqa hech qanday ma'lumot emas. Bu
# dasturchiga "boshliqda qaysi versiya ishlab turibdi" degan savolga
# javob berish uchun kerak (aks holda buni faqat so'rab bilish mumkin).
_PING_URL = "https://soddahisobot-telemetry.tasks-bot.workers.dev/ping"


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


# Ilova turli nomlar bilan tarqalgan: "SoddaHisobot.exe" (hozirgi) va
# "BankHisobotSoddalashtirish.exe" (eski). Foydalanuvchi qo'lda qayta
# nomlagan yoki Windows nusxa yaratgan bo'lishi ham mumkin. Shuning
# uchun nom BOSHIni emas, ICHIDAGI o'zakni qidiramiz — bu "Bank hisobot
# soddalashtirish (2).exe" kabi variantlarni ham tanib oladi.
# O'zaklar loyihaga xos so'zlardan iborat, shuning uchun boshqa dastur
# nomiga tasodifan to'g'ri kelib qolish ehtimoli yo'q.
_APP_FILE_STEMS = ("soddahisobot", "hisobotsoddalash")


def cleanup_old_versions():
    """Ilova joylashgan papkadagi eski nusxalarni va yangilanishdan
    qolgan vaqtinchalik fayllarni o'chiradi.

    Ilova ochilganda chaqiriladi. Nega kerak: foydalanuvchi bir necha
    marta qo'lda yuklab olsa, papkada "SoddaHisobot (1).exe",
    "SoddaHisobot (2).exe" kabi eski nusxalar to'planib qoladi va
    ish stoli yorlig'i eskisiga ishora qilib, "menda eski versiya
    ko'rinyapti" degan chalkashlik chiqadi.

    XAVFSIZLIK QOIDALARI (yangi nusxani xato o'chirib qo'ymaslik uchun):
      - faqat ishlab turgan .exe joylashgan PAPKADAGI fayllar;
      - faqat shu ilovaning nomiga o'xshash .exe fayllar;
      - faqat ishlab turgan .exe dan ESKIROQ fayllar (o'zgartirilgan
        vaqti bo'yicha) — yangiroq nusxaga tegilmaydi;
      - ishlab turgan faylning o'ziga hech qachon tegilmaydi.
    Xato chiqsa jim o'tkazib yuboriladi — bu shunchaki tozalash."""
    if not getattr(sys, "frozen", False):
        return
    try:
        exe_path = os.path.abspath(sys.executable)
        exe_dir = os.path.dirname(exe_path)
        exe_mtime = os.path.getmtime(exe_path)
    except Exception:
        return

    for name in os.listdir(exe_dir):
        full = os.path.join(exe_dir, name)
        low = name.lower()
        try:
            if not os.path.isfile(full) or os.path.samefile(full, exe_path):
                continue

            # a) yangilanishdan qolgan vaqtinchalik fayllar
            is_leftover = low.startswith(_OLD_EXE_PREFIX) or low.startswith("_update_")
            # b) shu ilovaning eskiroq nusxasi
            # Nomni normallashtiramiz: bo'shliq, chiziqcha, pastki chiziq va
            # qavslar olib tashlanadi. Shunda "Bank hisobot soddalashtirish
            # (2).exe", "SoddaHisobot - Copy.exe", "soddahisobot_1.exe" kabi
            # variantlar ham tanilib qoladi.
            flat = "".join(ch for ch in low if ch.isalnum())
            is_old_copy = (
                low.endswith(".exe")
                and any(stem in flat for stem in _APP_FILE_STEMS)
                and os.path.getmtime(full) < exe_mtime
            )
            if is_leftover or is_old_copy:
                os.remove(full)
        except Exception:
            continue  # band yoki ruxsat yo'q — tegmaymiz


def send_ping():
    """Ilova ochilganini va qaysi versiya ekanini xabar qiladi. To'liq
    "ovozsiz": internet yo'q bo'lsa yoki server javob bermasa, ilova
    ishlashiga umuman ta'sir qilmaydi. Alohida oqimda chaqirilishi kerak."""
    try:
        import platform

        payload = json.dumps({
            "host": platform.node() or "noma'lum",
            "version": APP_VERSION,
        }).encode("utf-8")
        req = urllib.request.Request(
            _PING_URL,
            data=payload,
            headers={"Content-Type": "application/json", "User-Agent": _USER_AGENT},
            method="POST",
        )
        urllib.request.urlopen(req, timeout=5).close()
    except Exception:
        pass

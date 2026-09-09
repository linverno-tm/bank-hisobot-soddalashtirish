# -*- coding: utf-8 -*-
"""
Nomlanmagan kontragentlar uchun Gemini yordamida guruh TAKLIF qilish.

MUHIM: bu modul hech qachon guruhni o'zi yozib qo'ymaydi — faqat taklif
beradi, oxirgi qarorni foydalanuvchi qabul qiladi. Buxgalteriyada
noto'g'ri kategoriya hisobotni jimgina buzadi va uni keyin topish qiyin,
shuning uchun "AI aytdi" degan asosda avtomatik yozish xavfli.

Tarmoq yo'q, kalit yo'q yoki javob buzuq bo'lsa — modul bo'sh natija
qaytaradi va ilova avvalgidek (qo'lda kiritish bilan) ishlayveradi.
"""
import json
import os
import sys
import urllib.request

MODEL = "gemini-3-flash-preview"
_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
_KEY_FILE = "ai_key.txt"

# Guruh nomlari qisqartma bo'lgani uchun AI ularni o'zicha tushunmaydi.
# Sinov shuni ko'rsatdi: izohsiz "Шахрихон туман МИБ" (Majburiy Ijro
# Byurosi) "МЕБ" (tovar) guruhiga qo'shib yuborilgan edi. Izoh qo'shilgach
# xato yo'qoldi, shuning uchun bu ro'yxat majburiy.
GROUP_HINTS = {
    "МЕБ": "tovar/mahsulot sotib olish (maishiy texnika, telefon, mototsikl, jihoz, xo'jalik mollari)",
    "Терминал": "HUMO / SmartVista terminal orqali tushum",
    "ф пайми": "moliya-hamkor (BNPL, nasiya) bilan hisob-kitob — TBC Fin Service, TBC BNPL",
    "ф вариант": "moliya-hamkor (BNPL, mikromoliya) — Variant Retail Finance, Uzum Nasiya",
    "банк хизмати": "bank komissiyasi, hisoblangan foizlar",
    "иш хаки ПК": "ish haqi, oylik to'lovi",
    "солик даромад": "jismoniy shaxs daromad solig'i",
    "солик КҚС": "qo'shilgan qiymat solig'i to'lovi",
    "солик фойда": "foyda solig'i to'lovi",
    "солик ижтимоий": "ijtimoiy soliq",
    "солик пенсия": "pensiya jamg'armasiga badal",
    "куриклаш": "qo'riqlash xizmati (IIB qoshidagi bo'lim)",
    "коммунал": "suv, gaz uchun to'lov",
    "электр": "elektr energiyasi uchun to'lov",
    "ижара": "ijara to'lovi",
    "хизмат": "xizmat ko'rsatish (IT, konsalting, dasturiy ta'minot, boshqa xizmatlar)",
    "СОРЖ": "SORJ",
}


def _base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def get_api_key():
    """Kalitni topadi: avval muhit o'zgaruvchisi, keyin ilova yonidagi
    ai_key.txt fayli.

    DIQQAT: kalitni .exe ichiga qo'shib bo'lmaydi — uni fayldan ajratib
    olish oson. Ommaga tarqatishdan oldin so'rovlar Cloudflare Worker
    orqali o'tkazilishi kerak, kalit esa faqat serverda qoladi."""
    key = os.environ.get("GEMINI_API_KEY", "").strip()
    if key:
        return key
    try:
        with open(os.path.join(_base_dir(), _KEY_FILE), encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return ""


def is_available():
    return bool(get_api_key())


def _build_prompt(items, groups):
    hints = {g: GROUP_HINTS[g] for g in groups if g in GROUP_HINTS}
    rows = [
        {"id": i, "nomi": str(it.get("name") or "")[:60], "matn": str(it.get("sample") or "")[:200]}
        for i, it in enumerate(items)
    ]
    return (
        "Sen O'zbekiston buxgalteriyasida bank ko'chirmalarini guruhlarga ajratasan.\n\n"
        f"MAVJUD GURUHLAR: {sorted(groups)}\n\n"
        f"GURUHLAR NIMANI ANGLATADI:\n{json.dumps(hints, ensure_ascii=False, indent=1)}\n\n"
        "QOIDALAR:\n"
        "- Faqat yuqoridagi ro'yxatdan tanla, yangi nom o'ylab topma.\n"
        "- O'xshash qisqartmalarni chalkashtirma: \"МИБ\" (Majburiy Ijro Byurosi) bu \"МЕБ\" EMAS.\n"
        "- Tovar sotib olish (texnika, transport, aloqa vositasi, xo'jalik mollari) -> МЕБ\n"
        "- Ishonching past bo'lsa yoki mos guruh bo'lmasa \"?\" yoz. Noto'g'ri taxmindan ko'ra \"?\" yaxshiroq.\n\n"
        "Javobni JSON massiv sifatida qaytar:\n"
        '[{"id":0,"guruh":"...","ishonch":"yuqori|past","sabab":"qisqa izoh"}]\n\n'
        f"Qatorlar:\n{json.dumps(rows, ensure_ascii=False, indent=1)}"
    )


def suggest(items, groups, timeout=60):
    """Har bir kontragent uchun guruh taklif qiladi.

    items  — [{"name": ..., "sample": ...}, ...] (find_unresolved natijasi)
    groups — ruxsat etilgan guruh nomlari to'plami

    Qaytaradi: {indeks: {"guruh": str, "ishonch": str, "sabab": str}}
    Xato yuz bersa — bo'sh lug'at (ilova avvalgidek ishlayveradi)."""
    key = get_api_key()
    if not key or not items or not groups:
        return {}
    try:
        body = json.dumps({
            "contents": [{"parts": [{"text": _build_prompt(items, groups)}]}],
            "generationConfig": {"responseMimeType": "application/json", "temperature": 0},
        }).encode("utf-8")
        req = urllib.request.Request(
            _API_URL.format(model=MODEL, key=key),
            data=body,
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        text = data["candidates"][0]["content"]["parts"][0]["text"]
        answers = json.loads(text)
    except Exception:
        return {}

    allowed = set(groups)
    out = {}
    for a in answers if isinstance(answers, list) else []:
        try:
            idx = int(a["id"])
            guruh = str(a.get("guruh") or "").strip()
        except (KeyError, TypeError, ValueError):
            continue
        # AI o'ylab topgan nomlarni qabul qilmaymiz — faqat mavjud guruhlar.
        if guruh not in allowed:
            continue
        out[idx] = {
            "guruh": guruh,
            "ishonch": str(a.get("ishonch") or "").strip(),
            "sabab": str(a.get("sabab") or "").strip(),
        }
    return out

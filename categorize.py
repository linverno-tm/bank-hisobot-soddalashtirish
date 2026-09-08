# -*- coding: utf-8 -*-
"""
Bank statement (Hamkorbank Client-Bank export) -> soddalashtirilgan variant.

Approach (per user's explicit warning): classify by the RAW account number
first (grouped), using payment-purpose text only to *propose* a category for
each unique raw account. This avoids misclassifying e.g. an auto/loan
payment that happens to also flow through HUMO/Uzkassa/SmartVista rails.
"""
import re
import json
import sys
from collections import defaultdict
from decimal import Decimal
import openpyxl

TOTAL_ROW_MARKERS = ("итоговый оборот", "итого", "jami", "всего")


def is_total_row(first_cell):
    if first_cell is None:
        return False
    return str(first_cell).strip().lower().startswith(TOTAL_ROW_MARKERS)


# Bitta hisob raqam (masalan G'aznachilik / Молия вазирлиги hisobvarag'i)
# orqali butunlay boshqa-boshqa maqsaddagi to'lovlar o'tadi: foyda solig'i,
# QQS, elektr uchun to'lov... Shuning uchun bu qoidalar hisob raqam bo'yicha
# o'rgatilgan guruhdan ham USTUN turadi — to'lov maqsadi (Назначение
# платежа) matni bu yerda hal qiluvchi hisoblanadi.
PURPOSE_FIRST_RULES = [
    (re.compile(r"фойда\s*соли[гғ]и", re.I), "солик фойда"),
    # DIQQAT: bu yerda faqat aniq "Кушилган киймат солиги" iborasi tekshiriladi.
    # Umumiy "НДС" so'zi ATAYLAB kiritilmagan — oddiy tovar to'lovlarida ham
    # "Сумма ... В т.ч. НДС (12%) ..." deb yoziladi, va u paytda bu soliq
    # to'lovi emas, balki narxning tarkibiy qismi. Umumiy "НДС" qoidasi
    # quyida, past darajali TEXT_RULES ichida qoldirilgan.
    (re.compile(r"[кқ]ушилган\s*[кқ]иймат\s*соли[гғ]и", re.I), "солик QQS"),
    (re.compile(r"ижтимоий\s*соли[кқ]", re.I), "солик ижтимоий"),
    (re.compile(r"даромадидан\s*олинадиган\s*соли[кқ]|даромад\s*соли[гғ]и", re.I), "солик даромад"),
    (re.compile(r"пенсия\s*бадалига", re.I), "солик пенсия"),
    (re.compile(r"сув\s*таъминоти|ичимлик\s*сув", re.I), "коммунал"),
    (re.compile(r"табиий\s*газ|газ\s*учун", re.I), "коммунал"),
    # Elektr alohida guruh — kommunalga qo'shilmaydi.
    (re.compile(r"фойдаланилган\s*электр|электр\s*учун|электр\s*энергия", re.I), "электр"),
    (re.compile(r"ижара\s*ту[лл]ови", re.I), "ижара"),
]


# Ordered keyword rules: (regex, category, confidence)
# confidence 'high' = auto-apply, 'guess' = apply but flag for review
TEXT_RULES = [
    (re.compile(r"начисленные\s*%%", re.I), "банк хизмати", "high"),  # naименование check done separately too
    (re.compile(r"smartvista|humo\s*\(|возмещение клиенту по покупкам тсп", re.I), "Терминал", "high"),
    (re.compile(r"tbc\s*fin\s*service|tbc\s*bnpl", re.I), "ф пайми", "high"),
    (re.compile(r"variant\s*retail\s*finance", re.I), "ф вариант", "high"),
    (re.compile(r"куриклаш", re.I), "куриклаш", "high"),
    (re.compile(r"зарплата|иш\s*ха[кқ]и", re.I), "иш хаки ПК", "high"),
    (re.compile(r"ижтимоий\s*соли[кқ]", re.I), "солик ижтимоий", "high"),
    (re.compile(r"даромадидан\s*олинадиган\s*соли[кқ]|даромад\s*соли[гғ]и", re.I), "солик даромад", "high"),
    (re.compile(r"пенсия\s*бадалига", re.I), "солик пенсия", "high"),
    (re.compile(r"сорж", re.I), "СОРЖ", "high"),
    # Lower confidence guesses (new patterns not seen in the DDD reference file yet).
    # Finance-partner style wording ("Публичная оферта" + "ген соглашение"-like BNPL
    # contracts) is checked BEFORE the generic "НДС" substring rule, since a BNPL
    # settlement text often mentions VAT only incidentally as a line item.
    (re.compile(r"оплата\s*100\s*%.*по\s*договору\s*публичная\s*оферта", re.I), "ф (aniqlanmagan)", "guess"),
    (re.compile(r"ижара\s*тулови|ижара\s*ту[лл]ови", re.I), "ижара", "guess"),
    (re.compile(r"фойдаланилган\s*электр|электр\s*учун", re.I), "электр", "guess"),
    (re.compile(r"консалтинг|konsalting", re.I), "хизмат", "guess"),
    (re.compile(r"фойда\s*соли[гғ]и", re.I), "солик фойда", "guess"),
    (re.compile(r"кушилган\s*[кқ]иймат\s*соли[гғ]и|\bндс\b", re.I), "солик QQS", "guess"),
]

GOODS_PURCHASE_HINT = re.compile(
    r"маиший техника|maishiy texnika|телефон|планшет|товар|жихоз|асбоб|"
    r"курилиш махсулот|дастурий таъминот|болалар уйинчок|мобил алока воситалари|"
    r"бытовой техник|бытавой техник|посуд",
    re.I,
)

# Vendor (INN) -> category, confirmed from a real "ДДД Август" file the boss
# already produced by hand. Same suppliers keep recurring across different
# companies' statements (same Andijon retail/appliance trade network), so
# this dictionary is the highest-confidence signal available and should grow
# over time (persist_dictionary.json).
_BUILTIN_VENDOR_INN = {
    "303389344": "МЕБ",       # Andijon Ravnaqi МЧЖ
    "207180749": "МЕБ",       # Goods And Services Impex
    "306776074": "ф пайми",   # TBC FIN SERVICE
    "312422124": "ф пайми",   # TBC BNPL
    "307490921": "ф вариант", # VARIANT RETAIL FINANCE
    "200292692": "солик пенсия",  # tuman ДСИ
    "200237592": "куриклаш",  # ИИБ КОШИДАГИ КУРИКЛАШ БУЛИМИ
    "310692639": "МЕБ",       # BIG ELECTRONICA МЧЖ
    "311019672": "МЕБ",       # ELEKTROMAX МЧЖ
    "309018562": "МЕБ",       # TOSHIBA AND MCHJ
    "310890749": "МЕБ",       # AMIR GROUP DU
}
KNOWN_VENDOR_INN = dict(_BUILTIN_VENDOR_INN)

import os
import sys


def _base_dir():
    # When bundled by PyInstaller, __file__ points into a temp extraction
    # folder, so use the actual .exe location instead to find/persist the
    # dictionary next to the app (editable, grows over time).
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


_DICT_PATH = os.path.join(_base_dir(), "persist_dictionary.json")

# Name-keyed learned mappings, for counterparties with no ИНН. Stored in the
# same persist_dictionary.json file, namespaced with a "NAME::" key prefix so
# it never collides with the (purely numeric) ИНН keys.
KNOWN_VENDOR_NAME = {}
_NAME_KEY_PREFIX = "NAME::"
# Foydalanuvchi o'zi qo'shgan "to'lov maqsadi matni -> guruh" qoidalari.
# Bitta hisob raqam orqali turli maqsaddagi to'lovlar o'tganda, ilovada
# tayyor qoida bo'lmasa, foydalanuvchi shu yerga o'zi qoida qo'sha oladi.
KNOWN_PURPOSE_TEXT = {}
_TEXT_KEY_PREFIX = "TEXT::"


def _load_persist_dict():
    try:
        with open(_DICT_PATH, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except FileNotFoundError:
        raw = {}
    for k, v in raw.items():
        if isinstance(k, str) and k.startswith(_TEXT_KEY_PREFIX):
            KNOWN_PURPOSE_TEXT[k[len(_TEXT_KEY_PREFIX):]] = v
        elif isinstance(k, str) and k.startswith(_NAME_KEY_PREFIX):
            KNOWN_VENDOR_NAME[k[len(_NAME_KEY_PREFIX):]] = v
        else:
            KNOWN_VENDOR_INN[k] = v


_load_persist_dict()


def save_learned_category(identifier, name, category):
    """Persist a user-supplied category for a previously-unresolved ("?")
    counterparty, so future reports auto-classify it. Keyed by Xisob raqam
    (Счет) when available (most reliable, and stable per counterparty);
    falls back to the exact counterparty name otherwise. `identifier` may
    also be an ИНН, for backward compatibility with entries learned by
    older versions of this app. Updates both the on-disk dictionary and the
    in-memory maps so the rest of the current run benefits immediately."""
    identifier = str(identifier).strip() if identifier else ""
    name = (name or "").strip()
    if not identifier and not name:
        return
    try:
        with open(_DICT_PATH, "r", encoding="utf-8") as f:
            raw = json.load(f)
    except FileNotFoundError:
        raw = {}
    if identifier.startswith(_TEXT_KEY_PREFIX):
        phrase = identifier[len(_TEXT_KEY_PREFIX):]
        raw[identifier] = category
        KNOWN_PURPOSE_TEXT[phrase] = category
    elif identifier:
        raw[identifier] = category
        KNOWN_VENDOR_INN[identifier] = category
    else:
        raw[f"{_NAME_KEY_PREFIX}{name}"] = category
        KNOWN_VENDOR_NAME[name] = category
    with open(_DICT_PATH, "w", encoding="utf-8") as f:
        json.dump(raw, f, ensure_ascii=False, indent=2, sort_keys=True)


def load_all_mappings():
    """Return the raw on-disk mapping ({ИНН or 'NAME::<nomi>': kategoriya}),
    for a UI to list/edit directly. Built-in (hardcoded) vendor mappings are
    NOT included here — only what the user has explicitly taught the app."""
    try:
        with open(_DICT_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def save_all_mappings(mapping):
    """Overwrite persist_dictionary.json wholesale with `mapping` (as
    returned/edited from load_all_mappings) and refresh the in-memory maps
    so a Guruhlar-manager UI's changes apply immediately, without discarding
    the built-in defaults."""
    with open(_DICT_PATH, "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2, sort_keys=True)
    KNOWN_VENDOR_INN.clear()
    KNOWN_VENDOR_INN.update(_BUILTIN_VENDOR_INN)
    KNOWN_VENDOR_NAME.clear()
    KNOWN_PURPOSE_TEXT.clear()
    for k, v in mapping.items():
        if isinstance(k, str) and k.startswith(_TEXT_KEY_PREFIX):
            KNOWN_PURPOSE_TEXT[k[len(_TEXT_KEY_PREFIX):]] = v
        elif isinstance(k, str) and k.startswith(_NAME_KEY_PREFIX):
            KNOWN_VENDOR_NAME[k[len(_NAME_KEY_PREFIX):]] = v
        else:
            KNOWN_VENDOR_INN[k] = v


def unique_counterparties(rows):
    """Return every unique counterparty found in a raw statement (keyed by
    Xisob raqam / Счет, or by name when the account number is missing),
    regardless of whether it was already classified. Used by the
    Guruhlar-manager UI to let the user pick a real counterparty out of an
    actual file (like browsing the statement itself) instead of typing an
    identifier by hand."""
    parties = {}
    for r in rows:
        account = str(r["account"]).strip() if r["account"] else ""
        inn = str(r["inn"]).strip() if r["inn"] else ""
        name = str(r["name"] or "").strip()
        if not account and not name:
            continue
        key = account if account else f"{_NAME_KEY_PREFIX}{name}"
        if key not in parties:
            parties[key] = {
                "account": account,
                "inn": inn,
                "name": name,
                "mfo": str(r.get("mfo") or "").strip(),
                "sample": str(r["purpose"] or "")[:200],
            }
    return parties


def find_unresolved(rows):
    """Scan rows and return unique unclassified ("?") counterparties, keyed
    by Xisob raqam / Счет (or by name when the account number is missing),
    for the caller to ask the user about before generating the report."""
    unresolved = {}
    for r in rows:
        cat, conf = classify_row(r["op"], r["name"], r["purpose"], r["inn"], r["account"])
        if conf != "review":
            continue
        account = str(r["account"]).strip() if r["account"] else ""
        name = str(r["name"] or "").strip()
        key = account if account else f"{_NAME_KEY_PREFIX}{name}"
        entry = unresolved.setdefault(key, {
            "account": account,
            "inn": str(r["inn"]).strip() if r["inn"] else "",
            "name": name,
            "mfo": str(r.get("mfo") or "").strip(),
            "sample": str(r["purpose"] or "")[:200],
            "count": 0,
        })
        entry["count"] += 1
    return unresolved


def classify_row(op, name, text, inn, account=None):
    """Classify a SINGLE row by its own text/name/Xisob raqam. Never looks
    at other rows sharing the same raw account number.

    Matching is done primarily by the counterparty's Xisob raqam (Счет).
    ИНН is kept as a secondary fallback so categories learned by older
    versions of this app (persist_dictionary.json entries keyed by ИНН)
    keep working after this update.

    ISTISNO: PURPOSE_FIRST_RULES — bitta hisob raqam orqali turli xil
    to'lovlar o'tadigan holatlar (G'aznachilik hisobvarag'i: foyda solig'i,
    QQS, elektr uchun to'lov) uchun to'lov maqsadi matni hisob raqamdan
    ustun turadi, aks holda hammasi bitta guruhga tushib qolardi."""
    name = name or ""
    text = text or ""

    if re.search(r"начисленные\s*%%", name, re.I):
        return "банк хизмати", "high"

    for rx, cat in PURPOSE_FIRST_RULES:
        if rx.search(text):
            return cat, "high"

    # Foydalanuvchi o'zi qo'shgan matn qoidalari — hisob raqamdan ustun
    # turadi, chunki ular aynan shunday "bitta hisob raqam, ko'p maqsad"
    # holatlarini qo'lda ajratish uchun kiritilgan.
    text_low = text.lower()
    for phrase, cat in KNOWN_PURPOSE_TEXT.items():
        if phrase.lower() in text_low:
            return cat, "high"

    account = str(account).strip() if account else ""
    inn = str(inn).strip() if inn else ""
    if account and account in KNOWN_VENDOR_INN:
        return KNOWN_VENDOR_INN[account], "high"
    if inn and inn in KNOWN_VENDOR_INN:
        return KNOWN_VENDOR_INN[inn], "high"

    name_key = name.strip()
    if name_key in KNOWN_VENDOR_NAME:
        return KNOWN_VENDOR_NAME[name_key], "high"

    for rx, cat, conf in TEXT_RULES:
        if rx.search(text) or rx.search(name):
            return cat, conf

    if op == 1 and GOODS_PURCHASE_HINT.search(text):
        return "МЕБ", "guess"

    return "?", "review"


def propose_category_for_group(op_values, names, texts, inns, accounts=None):
    """Given all rows sharing one raw account number, classify each row on
    its own merits, then only collapse to a single group-level category if
    every row agrees. A raw account is NOT trusted as a category by itself
    (e.g. a generic treasury account can carry profit tax, VAT AND utility
    payments; conversely two different accounts that both happen to mention
    HUMO/SmartVista text should NOT be assumed identical without agreeing)."""
    if accounts is None:
        accounts = [None] * len(op_values)
    per_row = [classify_row(op, n, t, i, a) for op, n, t, i, a in zip(op_values, names, texts, inns, accounts)]
    cats = {c for c, _ in per_row}
    if len(cats) == 1:
        cat, conf = per_row[0]
        return cat, conf, per_row
    return "ARALASH", "mixed", per_row


def load_raw_rows(path, sheet_name=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]
    header_row = None
    rows = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        vals = [c.value for c in row]
        if vals[0] == "Дата" and vals[1] == "Счет":
            header_row = row[0].row
            continue
        if header_row is None:
            continue
        if is_total_row(vals[0]):
            continue
        if all(v is None for v in vals):
            continue
        rows.append({
            "excel_row": row[0].row,
            "date": vals[0],
            "account": vals[1],
            "inn": vals[2],
            "name": vals[3],
            "doc_no": vals[4],
            "op": vals[5],
            "mfo": vals[6],
            "debit": vals[7],
            "credit": vals[8],
            "purpose": vals[9] or "",
        })
    return wb, ws, header_row, rows


def build_account_proposals(rows):
    groups = defaultdict(lambda: {"op": [], "names": [], "texts": [], "inns": [], "rows": [],
                                   "debit": Decimal(0), "credit": Decimal(0)})
    for r in rows:
        g = groups[r["account"]]
        g["op"].append(r["op"])
        g["names"].append(str(r["name"] or ""))
        g["texts"].append(str(r["purpose"] or ""))
        g["inns"].append(r["inn"])
        g["rows"].append(r)
        if r["debit"]:
            g["debit"] += Decimal(str(r["debit"]))
        if r["credit"]:
            g["credit"] += Decimal(str(r["credit"]))

    proposals = {}
    for acct, g in groups.items():
        accounts = [acct] * len(g["rows"])
        cat, conf, per_row = propose_category_for_group(g["op"], g["names"], g["texts"], g["inns"], accounts)
        proposals[acct] = {
            "category": cat,
            "confidence": conf,
            "count": len(g["rows"]),
            "debit_sum": g["debit"],
            "credit_sum": g["credit"],
            "sample_text": g["texts"][0][:160],
            "sample_name": g["names"][0],
            "rows": g["rows"],
            "per_row": per_row,
        }
    return proposals


if __name__ == "__main__":
    path = sys.argv[1]
    wb, ws, header_row, rows = load_raw_rows(path)
    proposals = build_account_proposals(rows)

    out_lines = []
    out_lines.append(f"Jami qatorlar: {len(rows)}, noyob hisob raqamlari: {len(proposals)}\n")
    for acct, info in sorted(proposals.items(), key=lambda kv: -kv[1]["count"]):
        if info["confidence"] == "mixed":
            out_lines.append(
                f"{acct!r:28} -> ARALASH (har qator alohida)  soni={info['count']:4}\n"
            )
            for r, (cat, conf) in zip(info["rows"], info["per_row"]):
                out_lines.append(
                    f"    [{cat:16}/{conf:6}] {r['name']} | {str(r['purpose'])[:140]}\n"
                )
            continue
        flag = "" if info["confidence"] == "high" else f"  <== TEKSHIRISH ({info['confidence']})"
        out_lines.append(
            f"{acct!r:28} -> {info['category']:20} soni={info['count']:4} "
            f"debet={info['debit_sum']:>15} kredit={info['credit_sum']:>15}{flag}\n"
            f"    namuna: {info['sample_name']} | {info['sample_text']}\n"
        )
    with open("account_proposals.txt", "w", encoding="utf-8") as f:
        f.writelines(out_lines)
    print("done, rows=", len(rows), "accounts=", len(proposals))

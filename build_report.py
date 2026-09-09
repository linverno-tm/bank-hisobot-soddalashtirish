# -*- coding: utf-8 -*-
"""
Xom Hamkorbank Client-Bank hisobotini (masalan "AB M Avgust.xlsx") boshliq
qo'lda tayyorlaydigan "soddalashtirilgan" formatga o'giradi:

  1. Har bir qatorning "Счет" ustuni (kontragent hisob raqami) kategoriya
     nomi bilan almashtiriladi (Терминал, банк хизмати, МЕБ, ...).
  2. "Лист1" degan yig'ma jadval varag'i qo'shiladi (SUMIF formulalar bilan,
     ДДД faylidagi ko'rinishga mos).

Ishlatilishi:
    python build_report.py "D:\\Telegram Desktop\\AB M  Avgust.xlsx" "natija.xlsx"

Bank hisobotlari bir necha xil ko'rinishda keladi (ustunlar tartibi va
sarlavha qatori har xil, ba'zilarida ИНН/Наименование ustunlari yo'q va
hisob raqam nom bilan bitta katakda turadi) — ustunlar nomi bo'yicha
aniqlanadi, categorize.load_raw_rows ga qarang.
"""
import os
import sys
import json
from decimal import Decimal
from copy import copy

import openpyxl
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

from categorize import load_raw_rows, classify_row, KNOWN_VENDOR_INN

DICT_PATH = "persist_dictionary.json"

# Yakuniy hisobotdagi yagona shrift (buyurtmachi so'roviga ko'ra)
FONT_NAME = "Times New Roman"
FONT_SIZE = 14


def _writable_cell(ws, row, col):
    """Katakka yozish uchun tayyor obyekt qaytaradi. Ba'zi hisobotlarda
    ustunlar birlashtirilgan (merged) bo'ladi — bunday katakka to'g'ridan
    to'g'ri yozib bo'lmaydi, faqat birlashma boshidagi katakka yoziladi.
    Shu boshlang'ich katakni topib beradi."""
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        return cell
    for rng in ws.merged_cells.ranges:
        if (rng.min_row <= row <= rng.max_row) and (rng.min_col <= col <= rng.max_col):
            return ws.cell(row=rng.min_row, column=rng.min_col)
    return None


def run(src_path, out_path):
    wb, ws, layout, rows = load_raw_rows(src_path)

    # classify every row independently, then sanity-check by raw account:
    # if one raw account ends up split across >1 category, surface it so it
    # gets extra attention (still applied per-row, never silently collapsed).
    from collections import defaultdict
    by_account = defaultdict(set)

    results = []  # (row_dict, category, confidence)
    for r in rows:
        cat, conf = classify_row(r["op"], r["name"], r["purpose"], r["inn"], r["account"])
        results.append((r, cat, conf))
        by_account[r["account"]].add(cat)

    # 1) "Счет" ustunini joyida kategoriya nomiga almashtiramiz (qolgan
    #    bezaklar, ustun kengliklari o'zgarmaydi). Ustun indeksi hisobot formatiga qarab har xil (B, D ...), shuning
    # uchun load_raw_rows aniqlagan joylashuvdan olamiz.
    account_col = layout["cols"]["account"] + 1  # openpyxl 1-asosli
    for r, cat, conf in results:
        cell = _writable_cell(ws, r["excel_row"], account_col)
        if cell is None:
            continue
        cell.value = cat
        cell.number_format = "@"

    # 2) build the Лист1 summary sheet (SUMIF over the report sheet)
    report_sheet_name = ws.title
    totals = defaultdict(lambda: [Decimal(0), Decimal(0)])  # cat -> [debet, kredit]
    for r, cat, conf in results:
        if r["debit"]:
            totals[cat][0] += Decimal(str(r["debit"]))
        if r["credit"]:
            totals[cat][1] += Decimal(str(r["credit"]))

    # Boshlang'ich/yakuniy qoldiq — "Остаток на начало/конец периода: ..."
    # matnli katakdan olinadi. Bu katak formatga qarab har xil joyda
    # turadi (A4 da yoki F4 da), shuning uchun ustun raqamiga tayanmasdan,
    # dastlabki qatorlar orasidan mos matnni qidiramiz.
    import re

    def parse_balance(s):
        m = re.search(r"([\d\s\xa0]+[.,]\d+|[\d\s\xa0]+)\s*$", s)
        if not m:
            return None
        # Bo'shliqlar (oddiy va uzilmas) — ming ajratuvchi, olib tashlanadi.
        # Vergul — kasr ajratuvchi, NUQTAGA aylantiriladi (O'CHIRILMAYDI —
        # avvalgi xato aynan shu yerda edi: vergulni butunlay o'chirib
        # tashlash butun sonni buzib yuborardi: "511 258,38" -> (eski)
        # "51125838" chiqar edi, to'g'risi 511258.38 bo'lishi kerak edi).
        t = m.group(1).replace(" ", "").replace("\xa0", "").replace(",", ".")
        try:
            return Decimal(t)
        except Exception:
            return None

    open_bal = close_bal = None
    for row in ws.iter_rows(min_row=1, max_row=min(10, ws.max_row)):
        for cell in row:
            text = str(cell.value or "")
            if not text:
                continue
            low = text.lower()
            if open_bal is None and "начало" in low:
                open_bal = parse_balance(text)
            if close_bal is None and "конец" in low:
                close_bal = parse_balance(text)
        if open_bal is not None and close_bal is not None:
            break

    if "Лист1" in wb.sheetnames:
        del wb["Лист1"]
    summary = wb.create_sheet("Лист1", 0)

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    acc_fmt = '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-'

    # Faqat fayl nomi ("АБ М Август"), to'liq yo'l emas. os.path.basename
    # ishlatiladi, chunki fayl tanlash oynasi yo'lni "/" bilan qaytaradi va
    # oldingi "\\" bo'yicha ajratish butun yo'lni sarlavhaga yozib qo'yardi.
    title = os.path.splitext(os.path.basename(src_path))[0]
    summary["B2"] = title
    summary["B2"].font = Font(bold=True)
    if open_bal is not None:
        summary["C2"] = float(open_bal)
    if close_bal is not None:
        summary["D2"] = float(close_bal)
    for col in ("B2", "C2", "D2"):
        summary[col].font = Font(bold=True)
        summary[col].border = border
        summary[col].number_format = acc_fmt

    headers = ["№", "Названия строк", " Дебет", " Кредит"]
    for i, h in enumerate(headers):
        c = summary.cell(row=3, column=1 + i, value=h)
        c.border = border
        c.number_format = acc_fmt

    row_i = 4
    total_debet = Decimal(0)
    total_kredit = Decimal(0)
    review_rows = []
    for idx, (cat, (deb, kred)) in enumerate(sorted(totals.items()), start=1):
        summary.cell(row=row_i, column=1, value=idx).border = border
        c_name = summary.cell(row=row_i, column=2, value=cat)
        c_name.border = border
        if deb:
            c = summary.cell(row=row_i, column=3, value=float(deb))
            c.border = border
            c.number_format = acc_fmt
            total_debet += deb
        if kred:
            c = summary.cell(row=row_i, column=4, value=float(kred))
            c.border = border
            c.number_format = acc_fmt
            total_kredit += kred
        for col in (1, 2, 3, 4):
            summary.cell(row=row_i, column=col).number_format = acc_fmt
            summary.cell(row=row_i, column=col).border = border
        row_i += 1

    summary.cell(row=row_i, column=2, value="Общий итог").border = border
    c = summary.cell(row=row_i, column=3, value=float(total_debet))
    c.number_format = acc_fmt
    c.border = border
    c = summary.cell(row=row_i, column=4, value=float(total_kredit))
    c.number_format = acc_fmt
    c.border = border
    summary.cell(row=row_i, column=1).border = border

    summary.column_dimensions["A"].width = 5.2
    summary.column_dimensions["B"].width = 24
    summary.column_dimensions["C"].width = 16
    summary.column_dimensions["D"].width = 16

    # Butun kitob bo'ylab yagona shrift: Times New Roman, 14. Qalinlik
    # (bold) va boshqa bezaklar qayerda bo'lsa, o'sha holicha saqlanadi.
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                old = cell.font
                cell.font = Font(
                    name=FONT_NAME,
                    size=FONT_SIZE,
                    bold=old.bold,
                    italic=old.italic,
                    underline=old.underline,
                    color=old.color,
                )

    wb.save(out_path)

    # 3) Past ishonch bilan belgilangan qatorlar soni. Avval bu ro'yxat
    #    alohida "_tekshirish.txt" fayliga yozilardi, lekin buyurtmachiga
    #    faqat Excel fayl kerak — endi bu son ilova jurnalida ko'rsatiladi.
    review_count = sum(1 for _r, _cat, conf in results if conf in ("guess", "review"))

    return {
        "out_path": out_path,
        "review_count": review_count,
        "total_debet": total_debet,
        "total_kredit": total_kredit,
    }


if __name__ == "__main__":
    info = run(sys.argv[1], sys.argv[2])
    # stdout may not accept Cyrillic under some Windows console codepages;
    # fall back to ascii-safe output rather than crashing after the file
    # (the one that matters) has already been saved successfully.
    try:
        print(f"Saqlandi: {info['out_path']}")
        print(f"Past ishonchli qatorlar: {info['review_count']} ta")
        print(f"Jami: debet={info['total_debet']} kredit={info['total_kredit']}")
    except UnicodeEncodeError:
        print("Saqlandi (nomi kirill belgilar tufayli konsolda ko'rsatilmadi).")

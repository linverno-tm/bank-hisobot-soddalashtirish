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

Natijada ikkinchi fayl (masalan natija_review.txt) — qaysi qatorlar past
ishonch bilan (guess) yoki umuman tekshirilmagan ("?") ekanini ko'rsatadi;
shularni Excelda ochib bir marta ko'zdan kechirish kifoya.
"""
import sys
import json
from decimal import Decimal
from copy import copy

import openpyxl
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter

from categorize import load_raw_rows, classify_row, KNOWN_VENDOR_INN

DICT_PATH = "persist_dictionary.json"


def run(src_path, out_path):
    wb, ws, header_row, rows = load_raw_rows(src_path)

    # classify every row independently, then sanity-check by raw account:
    # if one raw account ends up split across >1 category, surface it so it
    # gets extra attention in the review file (still applied per-row, never
    # silently collapsed).
    from collections import defaultdict
    by_account = defaultdict(set)

    results = []  # (row_dict, category, confidence)
    for r in rows:
        cat, conf = classify_row(r["op"], r["name"], r["purpose"], r["inn"])
        results.append((r, cat, conf))
        by_account[r["account"]].add(cat)

    # 1) rewrite the "Счет" column in place (column B), preserving all other
    #    formatting/merged cells/column widths from the original workbook
    col_letter = "B"
    for r, cat, conf in results:
        cell = ws[f"{col_letter}{r['excel_row']}"]
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

    # opening/closing balance, parsed from row 4 of the raw sheet
    a4 = str(ws["A4"].value or "")
    b4 = str(ws["B4"].value or "")

    def parse_balance(s):
        import re
        m = re.search(r"([\d\s.,]+)$", s)
        if not m:
            return None
        t = m.group(1).replace(" ", "").replace("\xa0", "").replace(",", "")
        try:
            return Decimal(t)
        except Exception:
            return None

    open_bal = parse_balance(a4)
    close_bal = parse_balance(b4)

    if "Лист1" in wb.sheetnames:
        del wb["Лист1"]
    summary = wb.create_sheet("Лист1", 0)

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    acc_fmt = '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-'

    title = src_path.split("\\")[-1].rsplit(".", 1)[0]
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

    wb.save(out_path)

    # 3) review list: rows that are guesses or unresolved, for a quick
    #    once-over in Excel before the file is treated as final
    review_lines = []
    for r, cat, conf in results:
        if conf in ("guess", "review"):
            review_lines.append(
                f"row {r['excel_row']:4} | {cat:18} ({conf:6}) | {r['name']} | {str(r['purpose'])[:130]}"
            )
    for acct, cats in by_account.items():
        if len(cats) > 1:
            review_lines.insert(0, f"OGOHLANTIRISH: hisob {acct} bir nechta toifaga bo'lindi: {cats}")

    review_path = out_path.rsplit(".", 1)[0] + "_tekshirish.txt"
    with open(review_path, "w", encoding="utf-8") as f:
        if review_lines:
            f.write("\n".join(review_lines))
        else:
            f.write("Hammasi yuqori ishonch bilan avtomatik belgilandi.\n")

    return {
        "out_path": out_path,
        "review_path": review_path,
        "review_count": len(review_lines),
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
        print(f"Tekshirish ro'yxati: {info['review_path']} ({info['review_count']} ta band)")
        print(f"Jami: debet={info['total_debet']} kredit={info['total_kredit']}")
    except UnicodeEncodeError:
        print("Saqlandi (nomi kirill belgilar tufayli konsolda ko'rsatilmadi).")

#!/usr/bin/env python3
"""Збирає третю таблицю: файл №2 із заповненим стовпцем B.

Для кожного серійника з A2, A3, … файлу №2 шукає такий самий серійник у стовпці A
файлу №1 і бере звідти B («Количество»), а якщо B порожній — C («Потеряно»).
Якщо серійника у файлі №1 немає, клітинка B лишається порожньою.
Файли №1 і №2 не змінюються; форматування файлу №2 переходить у результат.

Запуск:
    merge_serials.py 1.xlsx 2.xlsx            # результат — 3.xlsx поруч із 2.xlsx
    merge_serials.py 1.xlsx 2.xlsx out.xlsx
"""
import argparse
import sys
from pathlib import Path

try:
    import openpyxl
except ModuleNotFoundError:
    sys.exit("Потрібен openpyxl: pip install -r requirements.txt")


def key(value):
    """Приводить серійник до одного вигляду в обох файлах: 1111, 1111.0, ' 1111 ' → '1111'."""
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip() or None


def is_empty(value):
    return value is None or (isinstance(value, str) and not value.strip())


def load_lookup(path):
    """Серійник → значення з B, або з C, якщо B порожній. При дублях бере перший рядок."""
    ws = openpyxl.load_workbook(path, data_only=True).active
    lookup, dupes = {}, []
    for a, b, c in ws.iter_rows(min_row=2, min_col=1, max_col=3, values_only=True):
        k = key(a)
        if k is None:
            continue
        if k in lookup:
            dupes.append(k)
            continue
        lookup[k] = c if is_empty(b) else b
    return lookup, dupes


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("source", type=Path, help="файл №1: A — серійник, B — кількість, C — втрачено")
    p.add_argument("target", type=Path, help="файл №2: A — серійники, B буде заповнено")
    p.add_argument("output", type=Path, nargs="?", help="куди зберегти (типово 3.xlsx поруч із файлом №2)")
    p.add_argument("-f", "--force", action="store_true", help="перезаписати output, якщо він уже існує")
    args = p.parse_args()

    out = args.output or args.target.with_name("3.xlsx")
    if out.resolve() in (args.source.resolve(), args.target.resolve()):
        p.error("output не може збігатися з вхідним файлом")
    if out.exists() and not args.force:
        p.error(f"{out} уже існує — вкажіть інше ім'я або додайте --force")

    lookup, dupes = load_lookup(args.source)

    wb = openpyxl.load_workbook(args.target)
    ws = wb.active
    found, missing = 0, []
    for row in range(2, ws.max_row + 1):
        k = key(ws.cell(row, 1).value)
        if k is None:
            continue
        if k in lookup:
            ws.cell(row, 2).value = lookup[k]
            found += 1
        else:
            ws.cell(row, 2).value = None
            missing.append(k)
    wb.save(out)

    print(f"Готово: {out}")
    print(f"Знайдено: {found}, не знайдено: {len(missing)}")
    if missing:
        print("Немає у файлі №1:", ", ".join(missing))
    if dupes:
        print("Дублі у файлі №1 (узято перший рядок):", ", ".join(sorted(set(dupes))))


if __name__ == "__main__":
    main()

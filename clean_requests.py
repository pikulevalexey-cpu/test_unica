#!/usr/bin/env python3
"""
Приведение файла заявок к единому виду.

Вход:  Excel (.xlsx) или CSV/TXT с колонками Имя, Телефон, Дата заявки, Источник
Выход: один Excel-файл с двумя листами — "Чистые" и "Проблемные"

Запуск:
    python clean_requests.py <входной_файл.xlsx>
    python clean_requests.py <входной_файл.xlsx> --out-dir result --out-name результат.xlsx
"""

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

RU_MONTHS = {
    "январ": 1, "феврал": 2, "март": 3, "апрел": 4, "ма": 5, "июн": 6,
    "июл": 7, "август": 8, "сентябр": 9, "октябр": 10, "ноябр": 11, "декабр": 12,
}


def find_ru_month(word: str):
    word = word.lower()
    for stem, num in RU_MONTHS.items():
        if word.startswith(stem):
            return num
    return None


# ---------- Телефон ----------

def normalize_phone(raw):
    """Возвращает (нормализованный номер или None, ok: bool)."""
    if raw is None:
        return None, False
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none"):
        return None, False

    digits = re.sub(r"\D", "", s)

    if len(digits) == 11 and digits[0] in ("7", "8"):
        core = digits[1:]
    elif len(digits) == 10:
        core = digits
    else:
        return None, False

    if len(core) != 10:
        return None, False

    return f"+7{core}", True


# ---------- Дата ----------

def _try_build_date(year, month, day):
    try:
        y = int(year)
        if y < 100:
            y += 2000
        return datetime(y, int(month), int(day)).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return None


def normalize_date(raw):
    """Возвращает (ISO-дата или None, ok: bool)."""
    if raw is None:
        return None, False

    if isinstance(raw, datetime):
        return raw.strftime("%Y-%m-%d"), True

    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none"):
        return None, False

    # "13 марта 2026" / "04 апреля 2026"
    m = re.match(r"^(\d{1,2})\s+([а-яА-Я]+)\s+(\d{4})$", s)
    if m:
        day, month_word, year = m.groups()
        month = find_ru_month(month_word)
        if month:
            result = _try_build_date(year, month, day)
            if result:
                return result, True
        return None, False

    # "18 марта" — без года, восстановить нельзя
    m = re.match(r"^(\d{1,2})\s+([а-яА-Я]+)$", s)
    if m:
        return None, False

    # Числовые форматы: разделители . / - или пробел
    parts = re.split(r"[.\-/\s]+", s)
    parts = [p for p in parts if p != ""]

    if len(parts) != 3 or not all(p.isdigit() for p in parts):
        return None, False

    p0, p1, p2 = parts

    # yyyy-mm-dd / yyyy.mm.dd / yyyy/mm/dd
    if len(p0) == 4:
        result = _try_build_date(p0, p1, p2)
        return (result, True) if result else (None, False)

    # dd..mm..yyyy или dd..mm..yy (2-значный год)
    if len(p2) in (2, 4):
        # Основная гипотеза: день-первый (принято в РФ)
        result = _try_build_date(p2, p1, p0)
        if result:
            return result, True
        # Иначе пробуем месяц-первый (US mm/dd/yyyy)
        result = _try_build_date(p2, p0, p1)
        if result:
            return result, True

    return None, False


# ---------- Имя ----------

def normalize_name(raw):
    """Возвращает (имя или None, ok: bool)."""
    if raw is None:
        return None, False
    s = re.sub(r"\s+", " ", str(raw).strip())
    if not s or s.lower() in ("nan", "none"):
        return None, False

    def cap_token(tok):
        return "-".join(p[:1].upper() + p[1:].lower() if p else p for p in tok.split("-"))

    name = " ".join(cap_token(t) for t in s.split(" "))
    return name, True


# ---------- Основная обработка ----------

def process(df: pd.DataFrame):
    clean_rows = []
    problem_rows = []
    seen_phones = set()

    for _, row in df.iterrows():
        name, name_ok = normalize_name(row.get("Имя"))
        phone, phone_ok = normalize_phone(row.get("Телефон"))
        date, date_ok = normalize_date(row.get("Дата заявки"))
        source = row.get("Источник")
        source = None if pd.isna(source) else str(source).strip()

        reasons = []
        if not phone_ok:
            reasons.append("нет телефона")
        if not date_ok:
            reasons.append("битая дата")
        if not name_ok:
            reasons.append("нет имени")

        if reasons:
            problem_rows.append({
                "Имя (исходное)": row.get("Имя"),
                "Телефон (исходный)": row.get("Телефон"),
                "Дата заявки (исходная)": row.get("Дата заявки"),
                "Источник": source,
                "Проблема": "; ".join(reasons),
            })
            continue

        if phone in seen_phones:
            continue
        seen_phones.add(phone)

        clean_rows.append({
            "Имя": name,
            "Телефон": phone,
            "Дата заявки": date,
            "Источник": source,
        })

    clean_df = pd.DataFrame(clean_rows, columns=["Имя", "Телефон", "Дата заявки", "Источник"])
    problems_df = pd.DataFrame(
        problem_rows,
        columns=["Имя (исходное)", "Телефон (исходный)", "Дата заявки (исходная)", "Источник", "Проблема"],
    )
    return clean_df, problems_df


def load_input(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str)
    # csv / txt — пробуем угадать разделитель
    return pd.read_csv(path, dtype=str, sep=None, engine="python")


def save(clean_df: pd.DataFrame, problems_df: pd.DataFrame, path: Path):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        clean_df.to_excel(writer, sheet_name="Чистые", index=False)
        problems_df.to_excel(writer, sheet_name="Проблемные", index=False)


def main():
    parser = argparse.ArgumentParser(description="Приведение файла заявок к единому виду")
    parser.add_argument("input", help="Путь к файлу с заявками (.xlsx/.csv/.txt)")
    parser.add_argument("--out-dir", default=".", help="Папка для результата (по умолчанию текущая)")
    parser.add_argument("--out-name", default="заявки_результат.xlsx", help="Имя итогового файла (2 листа: Чистые / Проблемные)")
    args = parser.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        print(f"Файл не найден: {in_path}", file=sys.stderr)
        sys.exit(1)

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    df = load_input(in_path)
    clean_df, problems_df = process(df)

    out_path = out_dir / args.out_name
    save(clean_df, problems_df, out_path)

    print(f"Исходных строк: {len(df)}")
    print(f"Чистых уникальных записей (лист 'Чистые'): {len(clean_df)}")
    print(f"Проблемных строк (лист 'Проблемные'): {len(problems_df)}")
    print(f"Результат: {out_path}")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""pdf2excel.py – extrahiert Tabellen aus einem PDF, exportiert sie nach Excel
und sortiert ausgewählte Spalten nach vorne.

Abhängigkeiten: pdfplumber, pandas, openpyxl
Für ein einzelnes Binary kann anschließend pyinstaller verwendet werden.
"""

import argparse
from typing import List

import pdfplumber
import pandas as pd


def extract_tables(pdf_path: str) -> List[pd.DataFrame]:
    """Liest alle Tabellen aus einem PDF und gibt sie als DataFrames zurück."""
    frames: List[pd.DataFrame] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for raw_table in page.extract_tables():
                df = pd.DataFrame(raw_table[1:], columns=raw_table[0])
                frames.append(df)
    return frames


def export_excel(frames: List[pd.DataFrame], out_path: str, priority_cols: List[str]) -> None:
    """Speichert Tabellen in eine Excel-Datei und sortiert priority_cols nach vorne."""
    with pd.ExcelWriter(out_path) as writer:
        for idx, df in enumerate(frames, start=1):
            columns = priority_cols + [c for c in df.columns if c not in priority_cols]
            df[columns].to_excel(writer, sheet_name=f"Tabelle{idx}", index=False)


def main() -> None:
    parser = argparse.ArgumentParser(description="PDF → Excel mit Spalten-Umsortierung")
    parser.add_argument("pdf", help="Pfad zur PDF-Datei")
    parser.add_argument("xlsx", help="Zielpfad der Excel-Datei")
    parser.add_argument(
        "--move-first",
        nargs="+",
        default=[],
        help="Spaltennamen, die nach vorne sortiert werden sollen",
    )
    args = parser.parse_args()

    tables = extract_tables(args.pdf)
    export_excel(tables, args.xlsx, args.move_first)


if __name__ == "__main__":
    main()


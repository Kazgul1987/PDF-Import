#!/usr/bin/env python3
"""
Minimaler PDF→Excel-Konverter ausschließlich mit der Python-Standardbibliothek.

Einschränkungen:
- Funktioniert nur bei sehr einfachen, textbasierten PDFs.
- Tabellenerkennung basiert auf gleichbleibender Spaltenanzahl.
- Die erzeugte Excel-Datei enthält nur Basisdaten ohne Formatierung.
"""
import argparse
import re
import xml.etree.ElementTree as ET
from typing import List
from zipfile import ZipFile, ZIP_DEFLATED


def extract_lines(pdf_path: str) -> List[str]:
    """Entnimmt dem PDF naive Textzeilen."""
    with open(pdf_path, "rb") as f:
        data = f.read()
    matches = re.findall(rb"\(([^()]*)\)\s*T[Jj]", data)
    text = "\n".join(m.decode("latin1") for m in matches)
    return text.splitlines()


def detect_tables(lines: List[str]) -> List[List[List[str]]]:
    """Sucht Bereiche mit gleichbleibender Spaltenanzahl."""
    tables: List[List[List[str]]] = []
    current: List[List[str]] = []
    expected = None
    for line in lines + [""]:
        if not line.strip():
            if current:
                tables.append(current)
                current = []
                expected = None
            continue
        cols = line.strip().split()
        if expected is None or len(cols) == expected:
            current.append(cols)
            expected = len(cols)
        else:
            tables.append(current)
            current = [cols]
            expected = len(cols)
    return tables


def reorder_table(table: List[List[str]], priority: List[str]) -> List[List[str]]:
    if not table:
        return table
    header, *rows = table
    order = [header.index(c) for c in priority if c in header]
    order += [i for i in range(len(header)) if i not in order]
    def apply(row: List[str]) -> List[str]:
        return [row[i] if i < len(row) else "" for i in order]
    return [apply(header)] + [apply(r) for r in rows]


def write_xlsx(tables: List[List[List[str]]], out_path: str) -> None:
    def col_letter(idx: int) -> str:
        s = ""
        while idx >= 0:
            s = chr(ord("A") + idx % 26) + s
            idx = idx // 26 - 1
        return s

    with ZipFile(out_path, "w", ZIP_DEFLATED) as z:
        # [Content_Types].xml
        types = ET.Element("Types", xmlns="http://schemas.openxmlformats.org/package/2006/content-types")
        ET.SubElement(types, "Default", Extension="rels", ContentType="application/vnd.openxmlformats-package.relationships+xml")
        ET.SubElement(types, "Default", Extension="xml", ContentType="application/xml")
        ET.SubElement(types, "Override", PartName="/xl/workbook.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        for i in range(len(tables)):
            ET.SubElement(types, "Override", PartName=f"/xl/worksheets/sheet{i+1}.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
        z.writestr("[Content_Types].xml", ET.tostring(types, encoding="utf-8", xml_declaration=True))

        # _rels/.rels
        rels = ET.Element("Relationships", xmlns="http://schemas.openxmlformats.org/package/2006/relationships")
        ET.SubElement(rels, "Relationship", Id="rId1", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", Target="xl/workbook.xml")
        z.writestr("_rels/.rels", ET.tostring(rels, encoding="utf-8", xml_declaration=True))

        # xl/_rels/workbook.xml.rels
        wb_rels = ET.Element("Relationships", xmlns="http://schemas.openxmlformats.org/package/2006/relationships")
        for i in range(len(tables)):
            ET.SubElement(wb_rels, "Relationship", Id=f"rId{i+1}", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", Target=f"worksheets/sheet{i+1}.xml")
        z.writestr("xl/_rels/workbook.xml.rels", ET.tostring(wb_rels, encoding="utf-8", xml_declaration=True))

        # xl/workbook.xml
        workbook = ET.Element("workbook", xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main", attrib={"xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"})
        sheets = ET.SubElement(workbook, "sheets")
        for i in range(len(tables)):
            ET.SubElement(sheets, "sheet", name=f"Sheet{i+1}", sheetId=str(i+1), attrib={"{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id": f"rId{i+1}"})
        z.writestr("xl/workbook.xml", ET.tostring(workbook, encoding="utf-8", xml_declaration=True))

        # xl/worksheets/sheetX.xml
        for idx, table in enumerate(tables, start=1):
            ws = ET.Element("worksheet", xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            sheetData = ET.SubElement(ws, "sheetData")
            for r, row in enumerate(table, start=1):
                row_el = ET.SubElement(sheetData, "row", r=str(r))
                for c, value in enumerate(row):
                    cell = ET.SubElement(row_el, "c", r=f"{col_letter(c)}{r}", t="str")
                    ET.SubElement(cell, "v").text = value
            z.writestr(f"xl/worksheets/sheet{idx}.xml", ET.tostring(ws, encoding="utf-8", xml_declaration=True))


def main() -> None:
    parser = argparse.ArgumentParser(description="Einfacher PDF→Excel-Konverter ohne Zusatzmodule")
    parser.add_argument("pdf", help="Pfad zur PDF-Datei")
    parser.add_argument("xlsx", help="Zielpfad der Excel-Datei")
    parser.add_argument("--move-first", nargs="+", default=[], help="Spaltennamen, die nach vorn sortiert werden")
    args = parser.parse_args()

    lines = extract_lines(args.pdf)
    tables = detect_tables(lines)
    tables = [reorder_table(t, args.move_first) for t in tables]
    write_xlsx(tables, args.xlsx)


if __name__ == "__main__":
    main()

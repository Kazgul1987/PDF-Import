# PDF-Import

Ein kleines Tool, das Tabellen aus einer PDF-Datei extrahiert, als Excel-Datei speichert und ausgewählte Spalten nach vorne sortiert.

## Nutzung

```bash
python pdf2excel.py input.pdf output.xlsx --move-first Kohlenstoffgehalt
```

## Abhängigkeiten

```
pip install pdfplumber pandas openpyxl
```

## Ohne Installation

Mit [PyInstaller](https://pyinstaller.org/) kann ein einzelnes Binary erzeugt werden:

```bash
pip install pyinstaller
pyinstaller --onefile pdf2excel.py
```

Das entstehende Binary befindet sich im Verzeichnis `dist/` und kann ohne zusätzliche Installation ausgeführt werden.

## Reine Standardbibliothek (experimentell)

`pdf2excel_stdlib.py` kommt ohne Drittanbieter-Bibliotheken aus. Die Tabellenerkennung ist heuristisch und funktioniert nur bei sehr einfachen, textbasierten PDFs. Die erzeugte Excel-Datei enthält keine Formatierungen.

```bash
python pdf2excel_stdlib.py input.pdf output.xlsx --move-first Kohlenstoffgehalt
```


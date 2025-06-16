# gen-zeugnis

**gen-zeugnis** ist ein Python-Tool, um automatisch Schulzeugnisse im Word-Format (.docx) oder OpenDocument-Format (.odt) aus einer CSV-, XLSX- oder ODS-Datei mit Noten und einer Word-/ODT-Vorlage zu generieren. Es verwendet das Python-Modul [docxtpl](https://docxtpl.readthedocs.io/en/latest/) sowie pandas.

## Features

- Automatische Zeugnis-Generierung für beliebig viele Schüler:innen
- Noten und weitere Variablen werden aus einer CSV-, XLSX- oder ODS-Datei übernommen
- Flexible Word- (.docx) und ODT-Vorlagen mit Platzhaltern (z.B. `{{VN}}` für Vorname)
- Optional: Noten als Text (z.B. "sehr gut" statt "1"), aktivierbar mit `-mr`
- Fehlzeiten-Erfassung (Spalten: missed, excused, nonexcused)
- Ausgabe je Schüler:in als eigene Datei im Zielordner
- Automatische Datei- und Vorlagenerkennung im aktuellen Verzeichnis, falls keine Argumente übergeben werden

## Voraussetzungen

- Python 3.x
- [docxtpl](https://pypi.org/project/docxtpl/)
- [pandas](https://pypi.org/project/pandas/)
- [openpyxl](https://pypi.org/project/openpyxl/)
- [odfpy](https://pypi.org/project/odfpy/)

```bash
pip3 install docxtpl pandas openpyxl odfpy
```

## Nutzung

### 1. Vorlage erstellen

Lege eine Word- (`.docx`) oder ODT-Datei an, die Platzhalter im Format `{{VARNAME}}` enthält, z.B.:

```
Name: {{VN}} {{NN}}
Klasse: {{klasse}}
Mathematik: {{mathe}}
Deutsch: {{deutsch}}
...
```

### 2. Daten-Datei anlegen

Erstelle eine CSV-, XLSX- oder ODS-Datei mit den benötigten Spaltennamen als Header, passend zu den Platzhaltern. Beispiel für CSV:

```csv
VN,NN,klasse,mathe,deutsch,englisch,geburtsdatum,schuljahr,bemerkungen,missed,excused,nonexcused
Max,Mustermann,6a,1,2,2,01.01.2010,2024/2025,"Sehr engagiert.",2,1,1
```

### 3. Zeugnisse erzeugen

```bash
python gen-zeugnis.py schueler.csv vorlage.docx --outputfolder=zeugnisse
```

- Du kannst auch `.xlsx` oder `.ods` als Eingabedatei nutzen.
- Optional: Noten als Text ausgeben (`-mr`):

```bash
python gen-zeugnis.py schueler.csv vorlage.docx --outputfolder=zeugnisse -mr
```

**Hinweis:** Werden keine Datei-Argumente übergeben, versucht das Tool automatisch geeignete Daten- und Vorlagendateien im aktuellen Verzeichnis zu finden.

Jede/r Schüler:in erhält eine eigene Datei (benannt nach `NN` und `VN`) im angegebenen Ausgabeordner.

#### Alle Optionen

- `datafile` (Pfad zur Notenliste, auch `.xlsx` und `.ods` möglich; optional, automatische Suche)
- `templatefile` (Pfad zur Vorlage, `.docx` oder `.odt`; optional, automatische Suche)
- `--outputfolder` (Zielordner, Standard: `reports`)
- `-mr` oder `--marksreadable` (Noten als Text, Standard: Zahlen)

## Platzhalter

- Für jede Spalte in der Daten-Datei kann ein Platzhalter in der Vorlage verwendet werden (z.B. `{{mathe}}`).
- Standardmäßig werden Zahlen übernommen, mit `-mr` werden Noten in Textform eingesetzt ("sehr gut", "gut", ...).
- Zusätzliche Spalten z.B. für Bemerkungen, Datum, etc. sind möglich.
- Fehlzeiten können über die Spalten `missed`, `excused`, `nonexcused` erfasst und in der Vorlage platziert werden.

## Beispiel

Ein Beispiel für eine Zeugnissvorlage und eine passende CSV-Datei findest du oben oder kannst sie selbst anpassen.

## Windows EXE-Datei erzeugen

Mit [pyinstaller](https://www.pyinstaller.org/):

```bash
pyinstaller -F gen-zeugnis.py
```

---

**Autor:** Daniel Ache  
**Lizenz:** MIT
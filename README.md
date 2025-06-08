# gen-zeugnis

**gen-zeugnis** ist ein Python-Tool, um automatisch Schulzeugnisse im Word-Format (.docx) oder OpenDocument-Format (.odt) aus einer CSV-, XLSX- oder ODS-Datei mit Noten und einer Word-/ODT-Vorlage zu generieren. Es verwendet das Python-Modul [docxtpl](https://docxtpl.readthedocs.io/en/latest/) sowie pandas.

## Features

- Automatische Zeugnis-Generierung für beliebig viele Schüler:innen
- Noten und weitere Variablen werden aus einer CSV-, XLSX- oder ODS-Datei übernommen
- Flexible Word- (.docx) und ODT-Vorlagen mit Platzhaltern (z.B. `{{VN}}` für Vorname)
- Optional: Noten als Text (z.B. "sehr gut" statt "1"), aktivierbar mit `-mr 1`
- Fehlzeiten-Erfassung möglich (Spalten: missed, excused, nonexcused)
- Ausgabe je Schüler:in als eigene Datei im Zielordner

## Voraussetzungen

- Python 3.x
- [docxtpl](https://pypi.org/project/docxtpl/), [pandas](https://pypi.org/project/pandas/), [openpyxl](https://pypi.org/project/openpyxl/)

```bash
pip3 install docxtpl pandas openpyxl
```

## Nutzung

### 1. Vorlage erstellen

Lege eine Word- (`.docx`) oder ODT-Datei an, die Platzhalter im Format `{{VARNAME}}` enthält, zum Beispiel für Name, Klasse, Noten etc.:

```
Name: {{VN}} {{NN}}
Klasse: {{klasse}}
Mathematik: {{mathe}}
Deutsch: {{deutsch}}
...
```

### 2. Daten-Datei anlegen

Erstelle eine CSV-, XLSX- oder ODS-Datei mit den benötigten Spaltennamen als Header.
Die Spaltennamen müssen zu den Platzhaltern in deiner Vorlage passen. Beispiel für CSV:

```csv
VN,NN,klasse,mathe,deutsch,englisch,geburtsdatum,schuljahr,bemerkungen,missed,excused,nonexcused
Max,Mustermann,6a,1,2,2,01.01.2010,2024/2025,"Sehr engagiert.",2,1,1
```

### 3. Zeugnisse erzeugen

```bash
python gen-zeugnis.py schueler.csv vorlage.docx --outputfolder=zeugnisse
```

- Du kannst auch `.xlsx` oder `.ods` als Eingabedatei nutzen.
- Optional: Noten als Text ausgeben (`-mr 1`):

```bash
python gen-zeugnis.py schueler.csv vorlage.docx --outputfolder=zeugnisse -mr 1
```

Jede/r Schüler:in erhält eine eigene Datei (benannt nach `NN` und `VN`) im angegebenen Ausgabeordner.

#### Alle Optionen

- `csvfile` (Pfad zur Notenliste, auch `.xlsx` und `.ods` möglich)
- `docxfile` (Pfad zur Vorlage, `.docx` oder `.odt`)
- `--outputfolder` (Zielordner, Standard: `reports`)
- `-mr` oder `--marksreadable` (1 = Noten als Text, Standard: 0)

## Platzhalter

- Für jede Spalte in der Daten-Datei kann ein Platzhalter in der Vorlage verwendet werden (z.B. `{{mathe}}`).
- Standardmäßig werden Zahlen übernommen, mit `-mr 1` werden Noten in Textform eingesetzt ("sehr gut", "gut", ...).
- Zusätzliche Spalten z.B. für Bemerkungen, Datum, etc. sind möglich.
- Fehlzeiten können über die Spalten `missed`, `excused`, `nonexcused` erfasst und in der Vorlage platziert werden.

## Beispiel

Beispiel für eine Zeugnissvorlage und eine passende CSV findest du oben oder kannst sie selbst anpassen.

---

**Autor:** Daniel Ache  
**Lizenz:** MIT
# gen-zeugnis

**gen-zeugnis** ist ein Python-Tool, um automatisch Schulzeugnisse im Word-Format (.docx) aus einer CSV-Datei mit Noten und einer DOCX-Word-Vorlage zu generieren. Es verwendet das Python-Modul [docxtpl](https://docxtpl.readthedocs.io/en/latest/) zur Template-Verarbeitung.

## Features

- Automatische Zeugnis-Generierung für beliebig viele Schüler:innen
- Noten und weitere Variablen werden aus einer CSV-Datei übernommen
- Flexible Word-Vorlagen mit Platzhaltern (z.B. `{{VN}}` für Vorname)
- Optional: Noten als Text (z.B. "sehr gut" statt "1")
- Fehlzeiten-Erfassung möglich (spalten: missed, excused, nonexcused)

## Voraussetzungen

- Python 3.x
- [docxtpl](https://pypi.org/project/docxtpl/):  
  ```bash
  pip3 install docxtpl
  ```

## Nutzung

### 1. Vorlage erstellen

Lege eine Word-Datei (`.docx`) an, die Platzhalter im Format `{{VARNAME}}` enthält, zum Beispiel für Name, Klasse, Noten etc.:

```
Name: {{VN}} {{NN}}
Klasse: {{klasse}}
Mathematik: {{mathe}}
Deutsch: {{deutsch}}
...
```

### 2. CSV-Datei anlegen

Erstelle eine CSV-Datei mit den benötigten Spaltennamen als Header.
Die Spaltennamen müssen zu den Platzhaltern in deiner Vorlage passen. Beispiel:

```csv
VN,NN,klasse,mathe,deutsch,englisch,geburtsdatum,schuljahr,bemerkungen
Max,Mustermann,6a,1,2,2,01.01.2010,2024/2025,"Sehr engagiert."
```

### 3. Zeugnisse erzeugen

```bash
python gen-zeugnis.py schueler.csv vorlage.docx --outputfolder=zeugnisse
```

Optional: Noten als Text ausgeben (`-mr 1`):

```bash
python gen-zeugnis.py schueler.csv vorlage.docx --outputfolder=zeugnisse -mr 1
```

Jede/r Schüler:in erhält eine eigene DOCX-Datei im angegebenen Ausgabeordner.

## Platzhalter

- Für jede Spalte in der CSV kann ein Platzhalter in der Vorlage verwendet werden (z.B. `{{mathe}}`).
- Standardmäßig werden Zahlen übernommen, mit `-mr 1` werden Noten in Textform eingesetzt ("sehr gut", "gut", ...).
- Zusätzliche Spalten z.B. für Bemerkungen, Datum, etc. sind möglich.

## Beispiel

Beispiel für eine Zeugnissvorlage und eine passende CSV findest du oben oder kannst sie selbst anpassen.

---

**Autor:** Daniel Ache  
**Lizenz:** MIT  

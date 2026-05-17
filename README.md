# Energie-Tracker

Ein Python-Tool das tägliche Energie- und Gesundheitsdaten aus Markdown-Dateien in strukturierte Excel-Tabellen und interaktive Dashboards exportiert.

## Features

- 📊 **Markdown → Excel**: Konvertiert strukturierte Markdown-Logs zu Excel-Tabellen mit Formatierung
- 📈 **Automatische Analysen**: Wochenübersichten, Diagramme und Trendberechnung
- 🔍 **Korrelationsanalyse**: Zeigt welche Faktoren deinen Energiewert beeinflussen
- ⚡ **Einfache Integration**: Beliebige Markdown-Dateien, egal wo gespeichert

## 🚀 Geplante Features

- 🎨 **Dark-Mode HTML-Dashboard**: Interaktives Dashboard mit Catppuccin Mocha Theme
- 📊 Interaktive Diagramme mit plotly
- ⏱️ Konsum-Timeline mit Uhrzeiten
- 🔗 Korrelations-Heatmap
- 🎯 Filter & Zeitrahmen (7/30/Alle Tage)

## Unterstützte Datentypen

Das Tool parst folgende Metriken aus deinen Markdown-Logs:

- **Energie**: Energiewert (1-100), Stimmung
- **Schlaf**: Dauer, Puls, HRV, Atmung, Schlafphasen
- **Körper**: Gewicht, Körperfett, Muskelmasse, Körperwasser
- **Aktivität**: Schritte, aktive Zeit
- **Wasser/Konsum**: Holy Energy, Cannabis, Kaffee, Tee
- **Stress**: Subjektive Einschätzung

## Installation

### Voraussetzungen

- Python 3.8+
- pip (Python Package Manager)

### Setup mit virtueller Umgebung (empfohlen)

```bash
# 1. Repository klonen
git clone https://github.com/deinname/energie-tracker.git
cd energie-tracker

# 2. Virtuelle Umgebung erstellen
python3 -m venv venv

# 3. Aktivieren
source venv/bin/activate  # Linux/Mac
# oder
venv\Scripts\activate  # Windows

# 4. Abhängigkeiten installieren
pip install -r requirements.txt
```
### Schnell-Installation (ohne venv)

```bash
pip install openpyxl scipy 
```
### Verwendung
1. Markdown-Format vorbereiten

Deine energie_log.md sollte folgende Struktur haben:

    === ENERGIEWERT LOG === 
    Datum: DD.MM.JJJJ (Wochentag) 
    Energiewert: XX ("[Bezeichnung]")
    
    --- Schlaf --- 
    Schlafbeginn: HH:MM Uhr 
    Aufwachzeit: HH:MM Uhr 
    Schlafdauer gesamt: X Std XX Min 
    Schlafpuls: XX bpm 
    HRV: XX ms 
    Atmung: XX /min 
    Hauttemperatur: -X,X bis +X,X °C
    
    --- Schlafphasen --- 
    Wachphasen: XX Std XX Min 
    REM-Schlaf: XX Std XX Min 
    Leichtschlaf: XX Std XX Min 
    Tiefschlaf: XX Std XX Min
    
    --- Aktivität Vortag --- 
    Schritte: XXXXX 
    Aktive Zeit: XX Std XX Min 
    Art der Arbeit/Aktivität: [z.B. "viel laufen" / "LKW-Tag" / "frei"]
    
    --- Konsum Vortag --- 
    Holy Energy: [Anzahl Portionen + Uhrzeiten, z.B. "2x (06:00, 10:30)"]
    Kaffee/sonstiges Koffein: [Anzahl + Uhrzeiten, z.B. "1 Kaffee 05:30"] 
    Cannabis: [Menge + Uhrzeit, z.B. "0,25g um 22:00" / "kein"]
    
    --- Wasser Vortag --- 
    Holy Energy: [Anzahl x 500ml = XXXml] 
    Holy Hydration: [XXXml / "kein"] 
    Kaffee: [Anzahl x 150ml = XXXml / "kein"] 
    Tee: [Anzahl x XXXml = XXXml / "kein"] 
    Sonstiges: [XXXml / "kein"] 
    Gesamt: [XXXXml]
    
    --- Faktoren (Samsung) --- 
    Durchschnitt Schlafzeit: [Ausgezeichnet/Gut/Ausreichend/Achtung] 
    Regelmäßigkeit Schlafzeit: [Ausgezeichnet/Gut/Ausreichend/Achtung] 
    Schlafregelmäßigkeit: [Ausgezeichnet/Gut/Ausreichend/Achtung] 
    Schlafzeit: [Ausgezeichnet/Gut/Ausreichend/Achtung] 
    Aktivität Vortag: [Ausgezeichnet/Gut/Ausreichend/Achtung] 
    Aktivitätskontinuität: [Ausgezeichnet/Gut/Ausreichend/Achtung] 
    Schlafpuls: [Ausgezeichnet/Gut/Ausreichend/Achtung] 
    Schlaf-HRV: [Ausgezeichnet/Gut/Ausreichend/Achtung]
    
    --- Körper (nur montags) --- 
    Gewicht: XX,X kg 
    Körperfett: XX,X % 
    Muskelmasse: XX,X kg 
    Körperwasser: XX,X % 
    Skelettmasse: X,X kg

    --- Kontext --- 
    Stress (Watch): [Tagesschnitt: Entspannt/Niedrig/Mäßig/Hoch | Ausreißer: ...] 
    Stress (subjektiv): [Niedrig/Mittel/Hoch] 
    Besonderheiten: [...]
    
    ======================

Jeder Tag wird mit ====================== getrennt.
2. Skript anpassen

Öffne energie_tracker.py und passe die Pfade an:

    md_pfad    = "/pfad/zur/deiner/energie_log.md"
    excel_pfad = "/pfad/zum/ausgabe/energie_log.xlsx"

3. Ausführen

```bash
python3 energie_tracker.py
```

Ergebnis: energie_log.xlsx wird erstellt/aktualisiert mit drei Tabs:

    Rohdaten: Alle eingelesenen Werte
    Übersicht: Diagramme + Wochentrends
    Korrelationen: Welche Faktoren beeinflussen deine Energie?

### Output
Excel-Datei (energie_log.xlsx)
| Tab | Inhalt |
|---|---|
| **Rohdaten** | Alle Werte tabellarisch, sortiert nach Datum |
| **Übersicht** | 4 Diagramme + Hilfstabelle (Energie-Trend, Schlafphasen, HRV, Schritte) |
| **Korrelationen** | Pearson-Korrelation: Wie stark beeinflussen Schlaf, Wasser, Stress etc. deine Energie? |

### Technologie-Stack

- Python 3.8+
- openpyxl: Excel-Datei-Erstellung
- scipy: Korrelationsberechnung
- plotly (geplant): Interaktive HTML-Diagramme

### Roadmap
- [x] Excel-Export mit Formatierung
- [x] Korrelationsanalyse
- [ ] HTML-Dashboard (Phase 2)
- [ ] Interaktive Diagramme mit plotly
- [ ] Zeitraum-Filter (7/30/Alle Tage)
- [ ] Kennzahl-Karten (Durchschnitte, Trends)
- [ ] Konsum-Timeline mit Uhrzeiten
- [ ] Korrelations-Heatmap
- [ ] Catppuccin Mocha Dark Mode
- [ ] CSV-Export
- [ ] Datenvalidierung im Markdown
- [ ] Web-Interface (gehostete Version)
- [ ] Automatische Backups

### Lizenz

MIT License – siehe [LICENSE](LICENSE) Datei

### Contributing

Verbesserungsvorschläge, Bug-Reports und Pull Requests sind willkommen! 

Einfach ein Issue öffnen oder einen PR einreichen.

### Kontakt

Bei Fragen oder Feedback: [GitHub Issues](https://github.com/Sethur90/energie-tracker/issues)

### Made with ❤️ für Daten-Enthusiasten

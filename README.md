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
Schnell-Installation (ohne venv)

pip install openpyxl scipy 

Verwendung
1. Markdown-Format vorbereiten

Deine energie_log.md sollte folgende Struktur haben:

=== ENERGIEWERT LOG === 
Datum: 15.01.2025 (Dienstag) 
Energiewert: 85 ("[Ausgeschlafen]")

--- Schlaf --- 
Schlafbeginn: 23:30 Uhr 
Aufwachzeit: 07:30 Uhr 
Schlafdauer gesamt: 8 Std 00 Min 
Schlafpuls: 55 bpm 
HRV: 45 ms 
Atmung: 14 /min 
Hauttemperatur: -0,5 bis +0,5 °C

--- Schlafphasen --- 
Wachphasen: 0 Std 20 Min 
REM-Schlaf: 1 Std 30 Min 
Leichtschlaf: 3 Std 15 Min 
Tiefschlaf: 2 Std 55 Min

--- Aktivität Vortag --- 
Schritte: 8500 
Aktive Zeit: 45 Std 30 Min 
Art der Arbeit/Aktivität: [viel laufen]

--- Konsum Vortag --- 
Holy Energy: [2x (06:00, 10:30)]
Kaffee/sonstiges Koffein: [1 Kaffee 05:30] 
Cannabis: [0,25g um 22:00]

--- Wasser Vortag --- 
Holy Energy: [2 x 500ml = 1000ml] 
Holy Hydration: [500ml] 
Kaffee: [2 x 150ml = 300ml] 
Tee: [1 x 250ml = 250ml] 
Sonstiges: [kein] 
Gesamt: [2050ml]

--- Faktoren (Samsung) --- 
Durchschnitt Schlafzeit: [Ausgezeichnet] 
Regelmäßigkeit Schlafzeit: [Gut] 
Schlafregelmäßigkeit: [Gut] 
Schlafzeit: [Ausgezeichnet] 
Aktivität Vortag: [Ausreichend] 
Aktivitätskontinuität: [Gut] 
Schlafpuls: [Ausgezeichnet] 
Schlaf-HRV: [Gut]

--- Körper (nur montags) --- 
Gewicht: 75,2 kg 
Körperfett: 18,5 % 
Muskelmasse: 61,5 kg 
Körperwasser: 58,3 % 
Skelettmasse: 3,2 kg

--- Kontext --- 
Stress (Watch): [Tagesschnitt: Niedrig | Ausreißer: -] 
Stress (subjektiv): [Mittel] 
Besonderheiten: [Gut geschlafen, viel Bewegung]

======================

Jeder Tag wird mit ====================== getrennt.
2. Skript anpassen

Öffne energie_tracker.py und passe die Pfade an:

md_pfad    = "/pfad/zur/deiner/energie_log.md"
excel_pfad = "/pfad/zum/ausgabe/energie_log.xlsx"

3. Ausführen

python3 energie_tracker.py

Ergebnis: energie_log.xlsx wird erstellt/aktualisiert mit drei Tabs:

    Rohdaten: Alle eingelesenen Werte
    Übersicht: Diagramme + Wochentrends
    Korrelationen: Welche Faktoren beeinflussen deine Energie?

Output
Excel-Datei (energie_log.xlsx)
Tab 	Inhalt
Rohdaten 	Alle Werte tabellarisch, sortiert nach Datum
Übersicht 	4 Diagramme + Hilfstabelle (Energie-Trend, Schlafphasen, HRV, Schritte)
Korrelationen 	Pearson-Korrelation: Wie stark beeinflussen Schlaf, Wasser, Stress etc. deine Energie?
Technologie-Stack

    Python 3.8+
    openpyxl: Excel-Datei-Erstellung
    scipy: Korrelationsberechnung
    plotly (geplant): Interaktive HTML-Diagramme

Roadmap

    Excel-Export mit Formatierung
    Korrelationsanalyse
    HTML-Dashboard (Phase 2)
        Interaktive Diagramme mit plotly
        Zeitraum-Filter (7/30/Alle Tage)
        Kennzahl-Karten (Durchschnitte, Trends)
        Konsum-Timeline mit Uhrzeiten
        Korrelations-Heatmap
        Catppuccin Mocha Dark Mode
    CSV-Export
    Datenvalidierung im Markdown
    Web-Interface (gehostete Version)
    Automatische Backups

Lizenz

MIT License – siehe LICENSE Datei
Contributing

Verbesserungsvorschläge, Bug-Reports und Pull Requests sind willkommen! 

Einfach ein Issue öffnen oder einen PR einreichen.
Kontakt

Bei Fragen oder Feedback: GitHub Issues

Made with ❤️ für Daten-Enthusiasten

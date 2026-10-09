# Setup / Voraussetzungen

## 1. AutoIt3

- Download: https://www.autoitscript.com/site/autoit/downloads/
- Enthält **SciTE** (Editor zum direkten Ausführen, F5) und **Aut2Exe** (erzeugt eine `.exe`).
- Installation ist portabel nutzbar; ein voller Installer ist nicht zwingend.

Es werden **keine** zusätzlichen AutoIt-UDFs benötigt. Drag & Drop läuft über `DllCall("shell32.dll", ...)` —
die WinAPIEx-UDFs (`_WinAPI_RegisterDragDrop` usw.) sind **nicht** erforderlich und müssen nicht nachinstalliert werden.

## 2. PowerShell

- Auf Windows 7/Vista/8/10/11 **ab Werk vorhanden** (PowerShell 2.0 + .NET 3.5).
- Wird nur für die PDF-Funktionen genutzt:
  - **Extraktion** der eingebetteten Rechnungs-XML aus der PDF
  - **Erzeugung** der farbigen Sichtkopie und deren **Anhängen** an die PDF
- Verwendet ausschließlich Bordmittel (`System.IO.Compression.DeflateStream`, `System.Text.Encoding`) — **keine externen DLLs**.
- Reine **XML**-Dateien lassen sich auch ohne PowerShell verarbeiten.

## 3. Kompilieren / Starten

**Direkt starten (empfohlen zum Testen):**

1. `src/ZUGFeRD_Reader.au3` mit SciTE öffnen
2. **F5** drücken

**Als EXE bauen:**

1. Rechtsklick auf `src/ZUGFeRD_Reader.au3` → *Compile Script* (nutzt Aut2Exe)
2. Ergebnis: `ZUGFeRD_Reader.exe` (neben der `.au3`)

## 4. Erste Schritte

PDF ins Fenster ziehen → die Felder erscheinen automatisch → Button **🖨 Sichtkopie**.

## Hinweise

- Drag & Drop aus dem Explorer in ein **erhöhtes** (als Admin laufendes) Fenster wird von Windows blockiert.
  Dateien immer über den Button **📂 Öffnen** laden, wenn die App als Administrator läuft.
- Die **Original-PDF wird nie verändert** — die Sichtkopie wird immer in eine neue Datei geschrieben.
- Rein lesendes Tool: Es werden keinerlei Daten versendet.

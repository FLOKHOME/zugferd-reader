# ZUGFeRD Reader

**🇩🇪 Deutsch | [🇬🇧 English](README.en.md)**

Ein schlankes Windows-Tool (AutoIt3) zum **Einlesen, Prüfen und Aufbereiten** von E-Rechnungen im Format **ZUGFeRD / Factur-X / XRechnung**.

Rechnungs-PDF (oder XML) einfach ins Fenster ziehen — das Tool extrahiert die eingebettete XML, parst sie namensraum-unabhängig und zeigt Felder, Positionen und Summen. Auf Wunsch erzeugt es eine **farbige Sichtkopie** und hängt sie direkt an die Original-PDF an.

## Features

- **Drag & Drop** von PDF **oder** XML direkt ins Fenster (ohne Zusatz-UDFs, reine `DllCall`s)
- **Automatisch parsen** nach dem Laden
- **Format-Erkennung**: ZUGFeRD 1.x, ZUGFeRD 2.x / Factur-X (CII), XRechnung (UBL)
- **Namensraum-unabhängig** (XPath über `local-name()`) — funktioniert auch, wenn Absender eigene Präfixe verwenden
- **PDF-Anhang-Extraktion**: findet die eingebettete Rechnungs-XML (komprimiert oder unkomprimiert)
- **Anzeige**: Kopfdaten, Verkäufer/Käufer, Positionen, Summen, Prüfergebnis
- **Validierung**: Pflichtfelder, Währungscode, mathematische Prüfung (Netto + Steuer = Brutto)
- **Sichtkopie**: farbig aufbereitete PDF mit Kopfband, Zebra-Tabelle und Summen-Panel
- **Anhängen** der Sichtkopie an die Original-PDF (Incremental Update — das Original bleibt unverändert)
- **Export** der Felder als Text

## Voraussetzungen

- Windows (getestet auf Windows 7 x86; läuft auch auf Windows 10/11)
- [AutoIt3](https://www.autoitscript.com/site/autoit/downloads/) (SciTE zum Starten oder Aut2Exe zum Kompilieren)
- PowerShell 2.0+ (auf Windows 7 ab Werk vorhanden) — nur für die PDF-Extraktion/-Sichtkopie nötig

Details siehe [`setup/`](setup/README.md).

## Nutzung

1. `src/ZUGFeRD_Reader.au3` in SciTE öffnen und **F5** drücken (oder mit Aut2Exe eine `.exe` bauen).
2. Eine ZUGFeRD-/Factur-X-PDF (oder `.xml`) ins Fenster ziehen.
3. Die Felder erscheinen sofort. Reiter **Validierung** zeigt die Prüfergebnisse.
4. **🖨 Sichtkopie** erzeugt die farbige PDF — bei einer geladenen PDF wird sie automatisch **angehängt**.

## Wie es funktioniert

- ZUGFeRD/Factur-X ist ein **PDF/A-3** mit eingebetteter **CII-XML** (`factur-x.xml`).
- Die XML wird per PowerShell/.NET (`DeflateStream`) aus der PDF geholt: Streams werden gesucht, anhand des **eigenen Objekt-Dicts** als `/FlateDecode` erkannt (roh **und** entpackt probiert) und über ein echtes Tag (`<rsm:CrossIndustryInvoice` …) identifiziert.
- Geparst wird mit **MSXML** über `local-name()`-XPaths.
- Die Sichtkopie wird als PDF erzeugt (eigene Content-Streams) und per **Incremental Update** an die Original-PDF gehängt (neue Objekte + neue xref + `/Prev`); das Original wird nie überschrieben.

## Projektstruktur

```
src/ZUGFeRD_Reader.au3        AutoIt-GUI (Drag & Drop, Parsen, Sichtkopie)
ps/extract_pdf_xml.ps1        XML aus PDF extrahieren
ps/render_sichtkopie.ps1      farbige Sichtkopie erzeugen / anhängen
setup/README.md               Voraussetzungen & Kompilieren
```

## Lizenz

MIT — siehe [LICENSE](LICENSE).
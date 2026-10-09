# ZUGFeRD Reader

**[🇩🇪 Deutsch](README.md) | 🇬🇧 English**

A lightweight Windows tool (AutoIt3) to **read, validate and render** e-invoices in the **ZUGFeRD / Factur-X / XRechnung** format.

Drop an invoice PDF (or XML) onto the window — the tool extracts the embedded XML, parses it namespace-agnostically, and shows fields, line items and totals. Optionally it generates a **coloured visual copy** and appends it directly to the original PDF.

## Features

- **Drag & drop** of PDF **or** XML right onto the window (no extra UDFs, plain `DllCall`s)
- **Auto-parse** after loading
- **Format detection**: ZUGFeRD 1.x, ZUGFeRD 2.x / Factur-X (CII), XRechnung (UBL)
- **Namespace-agnostic** (`local-name()` XPath) — works even when senders use custom prefixes
- **PDF attachment extraction**: locates the embedded invoice XML (compressed or not)
- **Display**: header data, seller/buyer, line items, totals, validation result
- **Validation**: required fields, currency code, math check (net + tax = gross)
- **Visual copy (Sichtkopie)**: coloured PDF with a header band, zebra table and totals panel
- **Append** the visual copy to the original PDF (incremental update — the original stays untouched)
- **Export** the fields as text

## Requirements

- Windows (tested on Windows 7 x86; also runs on Windows 10/11)
- [AutoIt3](https://www.autoitscript.com/site/autoit/downloads/) (SciTE to run, or Aut2Exe to compile)
- PowerShell 2.0+ (ships with Windows 7) — only needed for PDF extraction / visual copy

See [`setup/`](setup/README.md).

## Usage

1. Open `src/ZUGFeRD_Reader.au3` in SciTE and press **F5** (or build an `.exe` with Aut2Exe).
2. Drag a ZUGFeRD/Factur-X PDF (or `.xml`) onto the window.
3. The fields appear immediately. The **Validation** tab shows the checks.
4. **🖨 Sichtkopie** creates the coloured PDF — for a loaded PDF it is **appended** automatically.

## How it works

- ZUGFeRD/Factur-X is a **PDF/A-3** with an embedded **CII XML** (`factur-x.xml`).
- The XML is pulled from the PDF via PowerShell/.NET (`DeflateStream`): streams are located, flagged as `/FlateDecode` using their **own object dict** (raw **and** inflated payloads are tried), and identified by a real tag (`<rsm:CrossIndustryInvoice` …).
- Parsing uses **MSXML** with `local-name()` XPaths.
- The visual copy is generated as a PDF (hand-built content streams) and appended via an **incremental update** (new objects + new xref + `/Prev`); the original is never overwritten.

## Layout

```
src/ZUGFeRD_Reader.au3        AutoIt GUI (drag & drop, parsing, visual copy)
ps/extract_pdf_xml.ps1        extract XML from PDF
ps/render_sichtkopie.ps1      render / append the coloured visual copy
setup/README.md               requirements & building
```

## License

MIT — see [LICENSE](LICENSE).
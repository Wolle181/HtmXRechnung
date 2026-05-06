# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel2ZUGFeRD converts German Excel invoices into ZUGFeRD-compliant PDF/XML invoices. It consists of:
- A compiled Python executable (`excel2zugferd.exe`) built with PyInstaller
- An Excel Add-in (`Excel2Zugferd.xlam`) with a custom ribbon button
- PowerShell scripts for building and installing the Add-in

The application is German-language throughout (UI, error messages, field labels).

## Build & Deployment

**Rebuild the Excel Add-in** (generates `C:\WORK\Excel2Zugferd.xlam`):
```powershell
powershell -ExecutionPolicy Bypass -File Create-Excel2Zugferd.ps1
```

**Install the Add-in** into Excel:
```powershell
powershell -ExecutionPolicy Bypass -File Install-Excel2Zugferd.ps1
# Or use the batch wrapper:
Install.bat
```

**Build the Python executable**: The `.exe` is a PyInstaller bundle. The Python source lives in `_internal/src/`. Rebuilding requires PyInstaller and the full Python environment (not currently configured in this repo).

There are no automated tests.

## Architecture

### Invocation Flow

The Excel ribbon button calls `RunMake()` in `vba_src/Excel2ZugferdMakro.bas`, which shells out to:
```
excel2zugferd.exe <sheet_number> "<excel_file_path>"
```

The executable supports two modes:
- **GUI mode**: No arguments → opens Tkinter window (`OberflaecheExcel2Zugferd`)
- **Batch/quiet mode**: Arguments provided → processes silently and exits

### Python Source (`_internal/src/`)

**Orchestration**:
- `middleware.py` — Central coordinator. Owns `IniFile`, `InvoiceCollection`, `ExcelContent`, PDF and ZUGFeRD handlers. Handles CLI args, error reporting, and both GUI/batch modes.

**Data Model**:
- `invoice.py` / `invoice_collection.py` — Invoice with all fields; collection assembles data from Excel
- `kunde.py`, `lieferant.py`, `konto.py`, `adresse.py` — Buyer, seller, bank account, address
- `steuerung.py` — Control parameters (create XML, GiroCode, Kleinunternehmer flag, filename pattern)

**Input**:
- `excel_content.py` — Reads `.xlsx` via pandas; normalizes columns to A/B/C; parses invoice number, date, line items, customer address
- `handle_ini_file.py` — Reads/writes company master data (INI format)
- `stammdaten.py` — Defines all master data field names (company info, bank account, column mappings)

**Output**:
- `handle_zugferd.py` — Builds ZUGFeRD XML using `drafthorse` (document class 380, conformance: Extended)
- `handle_pdf.py` — Generates PDF with `fpdf`; renders address window, line-item table, fold marks
- `handle_girocode.py` — Generates GiroCode QR for payment (IBAN/BIC/amount/reference)

**GUI (Tkinter)**:
- `oberflaeche_base.py` — Shared base window class
- `oberflaeche_excel2zugferd.py` — Main window: sheet list, "PDF erstellen" button
- `oberflaeche_ini.py` — Company master data form
- `oberflaeche_steuerung.py` — Control parameters form
- `oberflaeche_excelsteuerung.py` / `oberflaeche_excelpositions.py` — Column mapping forms

**Utilities**:
- `constants.py` — German error messages, unit mappings (e.g. `h→HUR`, `kg→KGM`)
- `windowseventlog.py` — Windows Event Log integration for error logging

### Key External Libraries (bundled in `_internal/`)
- `drafthorse` — ZUGFeRD XML generation
- `pandas` / `numpy` — Excel parsing
- `fpdf` — PDF generation
- `lxml` — XML/schema validation
- `Pillow` — Image handling for GiroCode and icons

### Configuration
Company master data and column mappings are stored in an INI file (path configurable at runtime). Field names are defined in `stammdaten.py`. The INI stores: company name, address, tax ID (Steuernummer/USt-IdNr), IBAN/BIC, and Excel column/row offsets for each invoice field.

### VBA Source (`vba_src/`)
- `Excel2ZugferdMakro.bas` — The ribbon button handler; calls `exe` and shows result in MsgBox
- `DebugTools.bas` — Utilities for exporting VBA modules during development

**IMPORTANT:** `vba_src/` is a **read-only, human-readable backup** only. It is NOT a build source.
The VBA code is embedded directly as a here-string inside `Create-Excel2Zugferd.ps1`.
When changing VBA logic, edit the here-string in the PS1 file — then update `vba_src/` manually as a mirror.
Never modify `Create-Excel2Zugferd.ps1` to read from `vba_src/` at build time.

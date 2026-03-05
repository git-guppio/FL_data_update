# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

FL_data_update is a Windows desktop application for batch-updating SAP Functional Location (FL) data via SAP GUI COM automation. It extracts FL codes from SAP, updates them to force cache refresh, and generates Excel reports with before/after field-level change tracking.

**Requirements:**
- Windows OS (COM automation is Windows-only)
- SAP GUI installed with scripting enabled (Options → Accessibility & Scripting → Scripting → Enable)
- At least one active SAP session logged in before running

## Running the Application

```bash
# Setup
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

# Run
python main.py
```

The GUI has a left panel for FL code input (one per line) and a right panel for real-time logs.

## Architecture

The project uses a layered architecture:

```
main.py (MainWindow/PyQt5)
    └── core/thread_manager.py (ThreadManager - concurrent.futures)
        └── sap/session_manager.py (SAPSessionManager - COM/win32com)
            └── SAP_Transactions.py (SAPDataExtractor - transaction logic)
                └── sap/utils.py (SAPUtils - clipboard, waits)

Support:
    core/logger.py     (ThreadSafeLogger - queue-based, non-blocking)
    core/base_component.py (BaseComponent mixin - unified self.log())
    config/settings.py (AppSettings - MAX_SAP_SESSIONS=4, timeouts, etc.)
```

**SAP_Transactions.py** is the primary transaction file (793 lines). The `sap/operations.py` file exists as a newer refactored version but `SAP_Transactions.py` is still the active implementation.

## Key Concepts

**FL Code Formats:**
- Exact: `ESS-ESND` (validated against regex `Mask_gen`)
- Wildcard: `ESS-ESSW*` (validated against `Mask_star`, uses SAP wildcard search)

**SAP Transactions Used:**
- `IH06` — Extract FL lists (with optional wildcard filtering)
- `SE16` — Data browser for table IFLO (FL details)
- `IL02` — FL modification (triggers SAP data cache refresh)

**Multi-language Support:**
SAP returns status text differently per language (IT/EN/ES/PT). Language-specific string mappings are in `SAP_Transactions.py` under `SAP_PARAMETERS` and in `main.py`'s `Check_Lang()` method.

**Data Flow:**
```
User Input → Validate (regex) → IH06 (extract FL list) →
SE16/IFLO (extract details) → IL02 (update to refresh) →
Compare before/after → Save Excel reports
```

**Threading Model:**
- `ThreadSafeLogger` uses `queue.Queue`; GUI polls it every 100ms via `QTimer`
- Each worker thread calls `pythoncom.CoInitialize()` / `CoUninitialize()` for COM safety
- `SAPSessionManager` distributes up to 4 sessions across threads
- Current branch (`Verifica-MultiThread`) is testing multi-threaded execution; single-thread version works

## Output Files

Generated in the working directory with timestamps:
- `FL_estratte_YYYYMMDD_HHMMSS.xlsx` — All extracted FL data
- `FL_aggiornate_YYYYMMDD_HHMMSS.xlsx` — Update results with change tracking

## Configuration

[config/settings.py](config/settings.py) — Key parameters:
- `MAX_SAP_SESSIONS = 4` — Max parallel SAP sessions
- `SAP_CONNECTION_INDEX = 0` — Which SAP connection to use if multiple open
- `SAP_TIMEOUT = 30` — Seconds to wait for SAP responses
- `DEFAULT_WORKERS = 4` — Thread pool size

## Important Implementation Details

- Clipboard is used as the data carrier between Python and SAP GUI (SAP scripting exports data via clipboard)
- `SAPSessionManager` uses `win32com.client` via the SAP Running Object Table (ROT)
- `BaseComponent` mixin provides `self.log()` to all major classes — always inherit from it in new SAP-related classes
- Column renaming in DataFrames must use `rename_columns_safely()` in `main.py` to handle duplicate headers from SAP clipboard exports
- The `check_modifications_detailed()` method in `main.py` tracks field-level changes for the Excel report

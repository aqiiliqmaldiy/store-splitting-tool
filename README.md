# Store Splitting Dashboard

A high-performance internal analytics dashboard for classifying, analyzing, and reporting store performance data across multiple reporting sets. Built for the **COA Team (R&D Rangers)**.

> [!IMPORTANT]
> **Multi-Country Auto-Detection:**  
> The tool automatically detects your target country based on the name of your folder path (supports **ID** - Indonesia, **TH** - Thailand, **BR** - Brunei, **MY** - Malaysia, **SG** - Singapore, **IN** - India). It will automatically map network drive files and configurations for the active region.

---

## Overview

The Store Splitting Dashboard allows analysts to:
- Load POSSYS and Store List data files
- Classify stores into up to 4 reporting sets (by Display Area, Store Type, and DA Slicer thresholds)
- Visualize distribution charts and sparklines
- Analyze Sales Performance by department
- Export comprehensive multi-sheet Excel reports
- Build custom pivot tables in an interactive drag-and-drop workspace

---

## Files

| File | Purpose |
|------|---------|
| `main.py` | Main application UI and window management (PyQt6) |
| `engine.py` | All data processing, math, and chart generation logic |

---

## Requirements

The app is self-healing — it will auto-install missing packages on startup.

- Python 3.9+
- `PyQt6`
- `pandas`
- `numpy`
- `scipy`
- `xlsxwriter`
- `python-calamine`
- `matplotlib`

---

## How to Run

```bash
python main.py
```

---

## Version History

---

### v3.5.17 — 2026-05-26
**Exported Excel Data Bars Restore**
- **In-Cell Progress Bars Restored:** Restored the missing conditional data bar charts (in-cell progress bars) in all summary sheets when exporting to `master.xlsx`.
- **Dynamic Header Color-Coding:** The bars now automatically scan column header texts, using the theme's emerald green for Sales, Turnover, and Contribution columns, and amber for Store Counts and Subtotals.

---

### v3.5.2 — 2026-05-20
**Regional Auto-Locate Bug Fix**
- **Dynamic network path reloading:** Switching countries in the dropdown now correctly re-scans the network drive to locate the selected country's Store List, Sales & Balance, and Grouping files, ensuring the correct data is used when updating the chart.
- **Dynamic active country profile label:** The UI now correctly displays the active country name on the right panel after a switch.

### v3.5.1 — 2026-05-20
**Thailand Store Count Display Fix**
- **Fixed active store count reconciliation:** Reconciled regional store counting to calculate unique active store codes directly from the master Store List sheet, showing exactly 1,276 total stores for Thailand (TH) and 1,361 total stores for Indonesia (ID).

---

### v3.5.0 — 2026-05-20
**Strict Network Start & Set 5 Ordering**
- **Reordered Set 5 reporting:** Subplots and tables are now arranged as SA < 8000, SA >= 8000, Mall < 10000, Mall >= 10000.
- **Terminated Scenario B (Offline Fallback):** Disabled local folder auto-detection at startup to enforce secure, strict network-only file scanning.
- **Dynamic Executable Path Sync:** Upgraded PyInstaller bootloader path resolution to correctly find the local Data directory right next to the .exe file when running as a compiled standalone application.
- **Dynamic Multi-Country Profile:** Replaced static Indonesia warnings in the README tab with a smart auto-detector (supporting ID, TH, BR, MY, SG, IN) that dynamically showcases your active country profile and syncs parameters accordingly.

---

### v3.4.7 — 2026-05-19
**Dark Green Bar Chart Palette Upgrade**
- **Rich Bar Color:** Updated primary bar charts to use a premium, deep dark green color (`#065F46` / deep pine emerald) for stronger impact and contrast.
- **Vibrant Accent Match:** Retained the dashboard's exact theme color `#10B981` (Emerald Green) for trend lines, numerical data labels, and sum table styling.

---

### v3.4.6 — 2026-05-19
**Emerald Green Custom Match**
- **Perfect Match:** Re-styled the primary charts to use the exact dashboard theme color `#10B981` (Emerald Green) for trend lines, numerical labels, and data tables.
- **Contrast & Legibility:** Selected a beautiful, soft pastel emerald `#A7F3D0` for all bar charts to deliver gorgeous contrast and premium readability.

---

### v3.4.5 — 2026-05-19
**Custom Sage Color Matching**
- **Precise Palette Match:** Styled the primary charts to use the user-requested `#B6CEB4` sage green for all bar charts and `#96A78D` dark sage green for trend lines, numerical texts, and layout tables.

---

### v3.4.4 — 2026-05-19
**Upgraded Primary Chart Color Palette**
- **Modern Curated Palette:** Upgraded primary plotting routines to use a premium, high-contrast palette (Indigo, Sky Blue, Emerald Green, and Amber) in place of the default Matplotlib `tab10` scheme for a clean, modern aesthetic.

---

### v3.4.3 — 2026-05-19
**Disabled Assist Tab**
- **Clean Workspace Optimization:** Disabled and hid the "Assist" tab from the main tab layout to focus the viewport on operational reporting.
- **Resource Conservation:** Deactivated dynamic background data science pipelines, multi-threading, and network file checks for the Assist tab.

---

### v3.4.2 — 2026-05-19
**Min and Max DA Columns on All Summary**
- **Detailed Size Boundaries:** Added "Min DA" and "Max DA" columns directly to the right of "Area" in the pivot tables on the All Summary tab.
- **Robust Multi-Column Syncing:** Seamlessly updated progress bar delegates, column widths, Grand Total highlights, and double-click drill-down actions to align with the new 8-column layout.

---

### v3.4.1 — 2026-05-19
**Dynamic Active Pivot Tables on All Summary**
- **Dynamic Table Hiding/Showing:** The All Summary tab now only displays tables that are active/picked in the side panel (primary selected set and compared sets).
- **Responsive Self-Scaling Layout:** Dynamically adjusts grid positions to beautifully fill the viewport when 1, 2, 3, or all 4 tables are active.

---

### v3.4.0 — 2026-05-19
**Instant Asynchronous Background Auto-Run**
- **Zero-Click Execution:** Configured the application to auto-trigger the size analysis immediately upon path sync inside `on_chart_ready`.
- **Silent Background Processing:** Leveraged non-blocking multi-threading (`QThread`) to run clustering optimizations silently in the background while the user works on other tabs.
- **Pre-Calculated Viewports:** Ensures that when you click on the "Assist" tab, the strategy tables, metrics charts, tier cards, and co-pilot are already loaded and waiting.

---

### v3.3.0 — 2026-05-19
**Gemini Data Analytic AI Co-Pilot**
- **Interactive Generative Agent:** Integrated real-time Generative AI chat interface as a dedicated 7th panel inside the "Assist" tab.
- **Data Science Context awareness:** Programmed standard prompts and customized chats to automatically ingest active metric averages, tier counts, and flagged outliers.
- **Secure Persistence:** Implemented local persistence for Gemini API key configuration to avoid repeat configurations across runs.
- **Multi-threaded Execution:** Powered queries using asynchronous QThreads to prevent GUI freezing.

---

### v3.2.0 — 2026-05-19
**Premium AI "Assist" Tab Upgrade**
- **Polymorphic Execution:** Resolved runtime crash by allowing `run_agent` to accept both raw string paths and standard DataFrames.
- **Dynamic Integration:** Automatically syncs paths and enables the "Assist" tab upon successful dashboard load in `on_chart_ready`.
- **Tier Card Correction:** Aligned metric summary cards with active High, Mid, and Low tier classifications.
- **Layer 0 Intelligence:** Added non-parametric Modified Z-score outlier/unicorn detection to locate exceptionally performing store flagships.

---

### v3.0.0 — 2026-04-30
**App Stability & Architecture Refactor**
- Optimized app stability and future-proofed the calculation logic for easier upgrades
- Cleaner background operations ensure the interface remains responsive during heavy tasks

---

### v2.4.0 — 2026-04-30
**Consolidated Multi-Sheet Reporting**
- Streamlined your reporting by exporting all analysis sets into a single, organized Excel file
- Future-proofed exports to automatically include any new store sets without requiring updates

---

### v2.3.1 — 2026-04-30
**Workflow Speed Tweaks**
- Faster workflow by removing unnecessary confirmation popups for empty tables
- Better organization by loading saved analyses directly into your active workspace tab
- Smarter data handling when switching between different source files ensures consistency

---

### v2.3.0 — 2026-04-29
**Workspace Persistence: Save & Load**
- Peace of mind: Save your custom pivot layouts and reload them anytime to pick up where you left off
- Reliable data restoration for complex statistical calculations and custom aggregations

---

### v2.2.2 — 2026-04-29
**Data Protection & UI Polish**
- Data integrity protection: Table cells are now locked to prevent accidental typing during analysis
- Clearer navigation with higher-visibility controls, better delete buttons, and hover states
- Improved readability with higher-contrast row selections

---

### v2.2.1 — 2026-04-29
**Enhanced Visual Hierarchy**
- Find your key results faster with distinct blue color-coding for all Grand Totals
- Bolding and high-contrast styling makes large summary tables easier to scan instantly

---

### v2.2.0 — 2026-04-28
**Interactive Pivot Workspace Launch**
- Total flexibility: A new drag-and-drop workspace to build unlimited custom reports side-by-side
- Simplified data grouping with a user-friendly multi-select filter popup

---

### v2.1.0 — 2026-04-28
**Sales Performance Analysis**
- Deep-dive into store efficiency with a new tab for detailed Sales and Turnover analysis
- Better categorization with integrated Dept and Store Set classification in one view

---

### v2.0.0 — 2026-04-27
**High-Performance Overhaul**
- Massive speed boost: Optimized calculations now handle large datasets instantly
- Eliminated app 'freezing' with smarter memory caching for repeated data reads

---

### v1.1.0 — 2026-04-23
**Configurable Set 4 Thresholds**
- More granular control over Set 4 reporting with a custom threshold (DA Slicer)
- Instant visual feedback when adjusting slicer values

---

### v1.0.1 — 2026-04-20
**Multi-Set & Quarter Analysis**
- Comprehensive analysis with support for 4 different store classifications simultaneously
- Automated trend tracking with new Quarter-over-Quarter analysis views

---

### v1.0.0 — 2026-04-17
**Initial Automation Release**
- Automated manual reporting: Raw POS data now generates charts and summaries in seconds
- Consolidated multiple store concepts into a single, unified view

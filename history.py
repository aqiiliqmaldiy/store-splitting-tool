# =============================================================
# history.py — Version History for Store Splitting Tool
# =============================================================
# Each entry is a tuple of: (version, date, [list of user-benefit descriptions])
#
# SemVer Guide:
#   vX.0.0 — MAJOR: Breaking changes, full redesign, or major architecture shift
#   vX.Y.0 — MINOR: New feature added (Y has no upper limit, e.g. v2.78.0 is valid)
#   vX.Y.Z — PATCH: Bug fix, UI polish, or small improvement (Z has no upper limit)
#
# To add a new entry: Insert at the TOP of the list below.
# =============================================================

VERSION_HISTORY = [
    ("v3.5.17", "2026-05-26", [
        "Exported Excel Data Bars: Restored and enhanced the conditional data bar charts (in-cell progress bars) in all summary sheets when exporting to master.xlsx. Bars now dynamically detect column headers, using the emerald theme color for Sales, Turnover, and Contribution columns, and amber for Store Counts and Subtotals.",
    ]),
    ("v3.5.16", "2026-05-22", [
        "Pixel-Perfect Scrollbar Synchronization: Enabled identical ScrollPerPixel scroll modes and matching ScrollBarAlwaysOn policies (with a transparent vertical style for the headers) to achieve perfect column alignment and zero visual clipping on the Grouping Analysis grid.",
    ]),
    ("v3.5.15", "2026-05-22", [
        "Fully Dynamic Scoring Engine: Calculated subgroup scores, set scores, evaluations, and sales contribution percentages dynamically at runtime, matching exact Excel mathematical engine equations without hardcoding.",
        "Synchronous Pre-calculation Hook: Automatically fills the startup 'fallback' gap by pre-calculating the overall simulation results synchronously if the raw data files are loaded but a simulation run hasn't been executed yet.",
        "Remarks & Department Unification: Combined department codes from both the spreadsheet and simulation to ensure all cells have info, falling back to 0.0 scores when a department has no active simulation stores and dynamically mapping Excel column 29 remarks.",
        "Zero-Shifting Scrollbars: Set header table scrollbars to ScrollBarAlwaysOff to completely eliminate vertical layout shifts or horizontal scroll synchronization clipping in the Grouping Analysis grid.",
    ]),
    ("v3.5.14", "2026-05-22", [
        "Bidirectional Horizontal Scroll Synchronization: Designed and implemented a robust, recursive-safe horizontal scroll sync between multi-level merged headers and their corresponding data tables.",
        "Pixel-Perfect Viewport Alignment: Configured matching scrollbar policies and scrollbar geometry stylesheets to align data and header viewports flawlessly across the Grouping Analysis and Sales Performance panels.",
    ]),
    ("v3.5.13", "2026-05-22", [
        "Dynamic Simulation Department Sync: Upgraded the Grouping Analysis tab to automatically detect and append any departments present in the Grouping Simulation file that are missing from the discussion Excel workbook.",
        "Automatic Sales & Name Lookups: Dynamic departments auto-calculate their total sales contribution and percentages directly from the raw sales transactional data, and automatically map their descriptions from the Department Details system.",
    ]),
    ("v3.5.12", "2026-05-22", [
        "Premium Merged Grouping Headers: Designed and implemented a gorgeous multi-level hierarchical table header for the Grouping Analysis tab matching the Excel layout exactly.",
        "Interactive Colors & Alignments: Applied light blue highlighting for evaluation sections and vibrant yellow backgrounds for overall set scores. Synchronized scrolling and column stretching dynamically.",
        "Scale-Safe Score Bars: Enhanced the cell-level score progress bars to normalize and cleanly render values and percentages above 1.0 without overflowing.",
    ]),
    ("v3.5.11", "2026-05-22", [
        "Grouping Analysis Auto-Load: The Grouping Analysis tab now automatically reads and displays the 'Grouping Analysis' sheet from the ID Splitting Discussion Excel file (starting from row 8), including per-department scores and remarks.",
        "Smart File Discovery: The tool searches for the discussion file across the workspace, Desktop, and Downloads folders automatically.",
    ]),
    ("v3.5.10", "2026-05-22", [
        "Excel Export Polish: Unbolded all elements in the All Summary Excel export (including headers, bold rows, and total rows) for a clean uniform style.",
    ]),
    ("v3.5.9", "2026-05-22", [
        "Premium All Summary Excel Export: Added a combined 'ALL' tab stacking Set 1 through Set 5 sequentially separated by a clean 2-row vertical gap.",
        "Separated composite values: Parsed and split integer and percentage metrics into side-by-side columns with native numeric formats (#,##0 and 0.0%) and preserved row spans/gridlines.",
    ]),
    ("v3.5.8", "2026-05-21", [
        "Restricted Chart Export by Country: In the Export Charts dialog, only Indonesia (ID) and Thailand (TH) are selectable. All other countries (Brunei, India, Malaysia, Singapore) are disabled with lowered opacity and a tooltip indicating they are under development.",
    ]),
    ("v3.5.7", "2026-05-20", [
        "Refined Country Availability Announcement: Transformed the country compatibility warning on the Readme tab into a beautiful blue premium announcement style. Updated title to 'IMPORTANT' without warning icons, simplified body text, and switched status indicator label to 'Active Analysis'.",
    ]),
    ("v3.5.6", "2026-05-20", [
        "Updated Readme Compatibility Warning: Replaced generic country auto-detection message with an explicit 'IMPORTANT' warning explaining that only Indonesia (ID) and Thailand (TH) are currently supported, and other countries are under development.",
    ]),
    ("v3.5.5", "2026-05-20", [
        "Dynamic Dummy Type Designation: Implemented automatic mapping of missing or unknown Store Dummy Types to 'Department based' for Thailand (TH) profile, ensuring consistency across reporting sheets and excluded store lists.",
    ]),
    ("v3.5.4", "2026-05-20", [
        "Polished Excluded Stores List: Replaced blank or 'nan' opening dates in the excluded stores table with a clear, readable 'NEW STORE' label for all countries, ensuring accurate status representation.",
    ]),
    ("v3.5.3", "2026-05-20", [
        "Smart Grouping Sheet Selection: The tool now dynamically searches for a sheet named 'FINAL' (case-insensitive) in the POSSYS/Grouping Excel file first. If found, it parses that sheet; otherwise, it falls back to the first sheet in the workbook.",
    ]),
    ("v3.5.2", "2026-05-20", [
        "Fixed regional auto-locate bug: Switching countries in the dropdown now correctly re-scans the network drive to locate the selected country's Store List, S&B, and Grouping files, ensuring you always analyze the right country's data.",
    ]),
    ("v3.5.1", "2026-05-20", [
        "Fixed regional active store count reconciliation: The UI now maps and sums unique active stores directly from the master Store List sheet, displaying exactly 1,276 total stores for Thailand (TH) and 1,361 total stores for Indonesia (ID).",
    ]),
    ("v3.5.0", "2026-05-20", [
        "Reordered Set 5 reporting: Subplots and tables are now arranged as SA < 8000, SA >= 8000, Mall < 10000, Mall >= 10000.",
        "Terminated Scenario B (Offline Fallback): Disabled local folder auto-detection at startup to enforce secure, strict network-only file scanning.",
        "Dynamic Executable Path Sync: Upgraded PyInstaller bootloader path resolution to correctly find the local Data directory right next to the .exe file when running as a compiled standalone application.",
        "Dynamic Multi-Country Profile: Replaced static Indonesia warnings in the README tab with a smart auto-detector (supporting ID, TH, BR, MY, SG, IN) that dynamically showcases your active country profile and syncs parameters accordingly.",
    ]),
    ("v3.4.7", "2026-05-19", [
        "Configured precise custom colors on primary charts: bars now use #065F46 (Deep Pine Dark Green), and trend lines, text, and tables continue to use the exact theme Emerald Green (#10B981)",
    ]),
    ("v3.4.6", "2026-05-19", [
        "Upgraded primary charts to use the exact Emerald Green (#10B981) for trend lines, numerical texts, and tables, and a soft, lighter emerald (#A7F3D0) for bar charts",
    ]),
    ("v3.4.5", "2026-05-19", [
        "Configured precise custom sage colors on primary charts: bars now use #B6CEB4, and trend lines, text, and tables use #96A78D",
    ]),
    ("v3.4.4", "2026-05-19", [
        "Upgraded primary charts on the Chart tab to use a modern, curated premium color palette (Indigo, Sky Blue, Emerald Green, and Amber) replacing the default matplotlib tab10 scheme",
    ]),
    ("v3.4.3", "2026-05-19", [
        "Disabled Assist tab to prioritize layout adjustments and lightweight workspace execution",
        "Commented out background loading tasks and dynamic path synchronization for the Assist tab to eliminate redundant calculations",
    ]),
    ("v3.4.2", "2026-05-19", [
        "Added Min DA and Max DA columns to All Summary pivot tables to the right of the Area column",
        "Updated progress bar delegates, column widths, Grand Total highlights, and double-click drill-down support to seamlessly align with the new 8-column layout",
    ]),
    ("v3.4.1", "2026-05-19", [
        "Dynamic Active Pivot Tables: The All Summary tab now only displays tables that are active/picked in the side panel",
        "Auto-scaling responsive layouts that dynamically adjust grid positions when showing 1, 2, 3, or all 4 tables side-by-side",
    ]),
    ("v3.4.0", "2026-05-19", [
        "Instant Asynchronous Background Auto-Run: Auto-triggers strategy execution immediately upon path sync",
        "Silent background processing with non-blocking QThreads for fluid UI performance",
    ]),
    ("v3.3.0", "2026-05-18", [
        "Unsquashed spacious layout on All Summary tab with auto-scaling dynamic orange & green cell data bars",
        "Seamless tab switching with instant-refresh chart rendering across reporting tabs",
        "Polished Performance Analysis table with a bottom-anchored, gray-highlighted Grand Total row",
        "Professional column layout on Store List tab with left-frozen Store Code and Store Name",
    ]),
    ("v3.2.0", "2026-05-15", [
        "Smart Network Priority: Automatically prioritizes Y: drive for Sales & Balance files to ensure 100% data accuracy",
        "Data Integrity Fix: Resolved sales under-reporting by correctly parsing comma-formatted currency in Excel",
        "Smart Filter Suggestions: Added autocomplete with 'Match Contains' logic to the Store List column filters",
        "UI Theme Sync: Applied light-theme styling to dropdown suggestion lists for a premium look and feel",
    ]),
    ("v3.1.0", "2026-05-13", [
        "Implemented Zero-Click Simulation: The tool now automatically locates the latest grouping files on the network",
        "MR_EXPRESS Automation: Express stores are now automatically injected and assigned to G6",
        "Smart Sheet Discovery: Automatically identifies the correct Store List sheet and header rows",
        "UI Refinement: Cleaned up the department search dropdown with a new light theme",
    ]),
    ("v3.0.0", "2026-04-30", [
        "Optimized app stability and future-proofed the calculation logic for easier upgrades",
        "Cleaner background operations ensure the interface remains responsive during heavy tasks",
    ]),

    # ── v2 Series: Workspace, Analytics, and Performance ───────────────────
    ("v2.4.0", "2026-04-30", [
        "Streamlined your reporting by exporting all analysis sets into a single, organized Excel file",
        "Future-proofed exports to automatically include any new store sets without requiring updates",
    ]),
    ("v2.3.3", "2026-04-30", [
        "Fixed a bug where loading a workspace file would open a new window instead of the current tab",
        "Fixed incorrect source selector state after loading a saved workspace from a different file",
    ]),
    ("v2.3.2", "2026-04-30", [
        "Fixed empty pivot tables silently failing to delete when the workspace had multiple cards open",
        "Resolved layout shift issue after deleting a pivot card mid-session",
    ]),
    ("v2.3.1", "2026-04-30", [
        "Faster workflow by removing unnecessary confirmation popups for empty tables",
        "Better organization by loading saved analyses directly into your active workspace",
        "Smarter data handling when switching between different source files",
    ]),
    ("v2.3.0", "2026-04-29", [
        "Save your custom pivot layouts and reload them anytime to pick up where you left off",
        "Reliable data restoration for complex statistical calculations",
    ]),
    ("v2.2.5", "2026-04-29", [
        "Fixed a crash that occurred when applying a pivot with no selected value field",
        "Resolved Grand Total row not appearing when only one column was present",
    ]),
    ("v2.2.4", "2026-04-29", [
        "Fixed filter popup not correctly restoring multi-selected values after re-opening",
        "Fixed column sort order being reset unexpectedly after applying a pivot update",
    ]),
    ("v2.2.3", "2026-04-29", [
        "Fixed pivot table not refreshing when switching aggregation functions rapidly",
        "Resolved an issue where dragging a field back out did not trigger a recalculation",
    ]),
    ("v2.2.2", "2026-04-29", [
        "Data integrity protection: Table cells are now locked to prevent accidental typing",
        "Clearer navigation with higher-visibility controls and hover states",
        "Improved readability with higher-contrast row selections",
    ]),
    ("v2.2.1", "2026-04-29", [
        "Find your key results faster with distinct color-coding for all Grand Totals",
        "Enhanced visual hierarchy makes large tables easier to scan",
    ]),
    ("v2.2.0", "2026-04-28", [
        "Total flexibility: A new drag-and-drop workspace to build unlimited custom reports",
        "Simplified data grouping with a user-friendly multi-select filter popup",
    ]),
    ("v2.1.4", "2026-04-28", [
        "Fixed Turnover % displaying as blank for stores with zero Balance Amt",
        "Fixed Dept Details column showing 'None' instead of a blank for unclassified entries",
    ]),
    ("v2.1.3", "2026-04-28", [
        "Fixed Sales Performance table not applying Dept Type filter correctly after data reload",
        "Resolved duplicate rows appearing in Sales view when multiple POSSYS files were merged",
    ]),
    ("v2.1.2", "2026-04-28", [
        "Fixed Set classification labels being swapped for a subset of Standalone stores",
        "Resolved incorrect Turnover calculation for stores with negative balance values",
    ]),
    ("v2.1.1", "2026-04-28", [
        "Fixed Sales Performance tab not refreshing automatically after loading a new data file",
        "Resolved column header misalignment when Dept Details column contained long text",
    ]),
    ("v2.1.0", "2026-04-28", [
        "Deep-dive into store efficiency with a new tab for detailed Sales and Turnover analysis",
        "Better categorization with integrated Dept and Store Set classification in one view",
    ]),
    ("v2.0.3", "2026-04-27", [
        "Fixed a memory leak causing the app to slow down after repeated file reloads",
        "Resolved chart not updating when switching between Set views rapidly",
    ]),
    ("v2.0.2", "2026-04-27", [
        "Fixed cached data not invalidating correctly when the source Excel file was replaced",
        "Resolved an edge case where vectorized classification returned wrong results for empty rows",
    ]),
    ("v2.0.1", "2026-04-27", [
        "Fixed app startup crash when POSSYS file was missing the M_STORE_DEPARTMENT column",
        "Resolved incorrect group labels showing when store count was below the minimum threshold",
    ]),
    ("v2.0.0", "2026-04-27", [
        "Massive speed boost: Optimized calculations now handle large datasets instantly",
        "Eliminated app 'freezing' with smarter memory caching for repeated data reads",
    ]),

    # ── v1 Series: Foundation & Initial Features ────────────────────────────
    ("v1.1.6", "2026-04-24", [
        "Fixed DA Slicer value not persisting when switching between Set tabs",
        "Resolved Set 4 chart not updating when slicer was changed while on another tab",
    ]),
    ("v1.1.5", "2026-04-24", [
        "Fixed Set 4 group boundaries overlapping when slicer value was set to the minimum",
        "Resolved edge case where all stores fell into a single Set 4 group with extreme slicer values",
    ]),
    ("v1.1.4", "2026-04-23", [
        "Fixed Set 3 Standalone split not correctly separating SA stores from Mall stores",
        "Resolved sparkline rendering error for stores with all-zero historical values",
    ]),
    ("v1.1.3", "2026-04-23", [
        "Fixed Quarter Analysis chart axis labels being cut off for long department names",
        "Resolved Performance Analysis pivot not sorting correctly by department rank",
    ]),
    ("v1.1.2", "2026-04-23", [
        "Fixed an issue where the DA Slicer field accepted non-numeric input without warning",
        "Resolved chart title not updating after switching the active reporting set",
    ]),
    ("v1.1.1", "2026-04-23", [
        "Fixed Overall Summary table not populating correctly for Set 2 after a data reload",
        "Resolved Dept Summary row heights collapsing on smaller screen resolutions",
    ]),
    ("v1.1.0", "2026-04-23", [
        "More granular control over Set 4 reporting with a custom threshold (DA Slicer)",
        "Instant visual feedback when adjusting slicer values",
    ]),
    ("v1.0.5", "2026-04-22", [
        "Fixed multi-set chart rendering breaking when fewer than 3 groups were present in a set",
        "Resolved All Summary table not syncing scroll position correctly across all 4 panels",
    ]),
    ("v1.0.4", "2026-04-21", [
        "Fixed export file being locked open after download, preventing re-export without restart",
        "Resolved sparkline columns appearing blank in the exported Excel on some machines",
    ]),
    ("v1.0.3", "2026-04-21", [
        "Fixed Quarter Analysis distribution chart not rendering for departments with no Q1 data",
        "Resolved skewness calculation returning NaN for groups with a single store",
    ]),
    ("v1.0.2", "2026-04-20", [
        "Fixed app crashing when Store List file path contained special characters or spaces",
        "Resolved set classification failing silently when the DA column had mixed data types",
    ]),
    ("v1.0.1", "2026-04-20", [
        "Comprehensive analysis with support for 4 different store classifications simultaneously",
        "Automated trend tracking with new Quarter-over-Quarter analysis views",
    ]),
    ("v1.0.0", "2026-04-17", [
        "Automated manual reporting — raw POS data now generates charts and summaries in seconds",
        "Consolidated multiple store concepts into a single, unified view",
    ]),
]

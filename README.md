# 🧭 Bosla Toolbox for ArcGIS Pro

**Bosla (بوصلة) — Arabic for "compass." Because every survey starts with one.**

![ArcGIS Pro](https://img.shields.io/badge/ArcGIS%20Pro-3.x-blue)
![Python](https://img.shields.io/badge/Python-3.x-green)
![License](https://img.shields.io/badge/License-MIT-yellow)
![Version](https://img.shields.io/badge/Version-1.3.0-orange)

Field crews, surveyors, and project stakeholders live in Excel. Your data lives in ArcGIS Pro. Permits need coordinates in DMS. Databases need bulk edits nobody dares to make. Bosla bridges these gaps: **8 production-grade tools** for the Excel round trip, coordinate conversion in both directions, data quality control, and safe database maintenance — built by a geomatics professional for daily production work in surveying and upstream oil & gas.

> ### ⚠️ Built for production data
> Every destructive operation in Bosla is designed to be safe by default:
> - **Apply Excel Edits** runs in **dry-run mode by default** — it reports every change *without writing anything* until you explicitly uncheck dry-run.
> - An **automatic timestamped backup** is created before any update.
> - Updates run inside an **edit session** where possible (versioned enterprise geodatabases supported), with rollback on failure.
> - Invalid values (wrong type, over-length text, unparseable dates) are **skipped and reported per cell** — never silently written, never silently truncated.
> - The **Backup Manager** restores only to *new* datasets — it never overwrites existing data.

---

## The 8 Tools

### 01 — Export Attributes To Excel
Export any layer or table's attributes to a formatted, filter-ready .xlsx.

- Styled header, frozen panes, auto-filter, sized columns
- Key field tracked automatically (defaults to ObjectID) for safe re-import
- Hidden metadata sheet records the source dataset, key field, and toolbox version — Tool 02 reads it back to protect you from applying edits to the wrong layer
- Supports feature classes and standalone tables

**Use it for:** bulk editing in Excel, QA/QC review, business reporting, data exchange with non-GIS colleagues.

### 02 — Apply Excel Edits Back To Layer
Apply the edits made in Excel directly back to your ArcGIS Pro layer — safely.

- Key-field matching with automatic type normalization (no more silent "0 rows matched" when Excel turns `1` into `1.0`)
- Every value validated and cast to the target field's type before writing
- **Dry-run by default** · automatic backup · edit-session support for enterprise geodatabases
- Warns if the Excel file was exported from a different layer or key field
- Selective field updates, blank-cell handling options, per-cell error reporting

**Use it for:** mass attribute corrections, database maintenance, vendor data integration, QA/QC fix cycles.

### 03 — Outline To Vertices Coordinates (Universal CS)
Convert polygon/polyline vertices to points and export coordinates in **any** coordinate system ArcGIS Pro supports.

- Projected (X/Y) and geographic outputs in the same run
- Geographic formats: Decimal Degrees, **DMS** (`31° 07' 24.441" E`), and **DDM**
- **Datum transformation support** — pick lists auto-populate for your input/output pair, and the tool *warns you* when datums differ and no transformation is selected (the silent error that shifts Egypt 1907 ↔ WGS84 coordinates by tens of meters)
- Universal coordinate system picker: WGS84, UTM zones, Egypt 1907, national grids, engineering systems, custom PRJ — no code changes needed
- Multipart features, Z values, closing-vertex control, projected once per feature for speed on dense boundaries

**Use it for:** survey coordinate reports, permit documents, concession boundary listings, control point verification.

### 04 — Import Points From Excel 🆕
The reverse direction: turn any coordinate list into a point feature class.

- Parses **Decimal Degrees, DMS, DDM, and projected XY** — with auto-detect
- Accepts real-world messy formats: `31° 07' 24.441" E`, `N 29 59 12.645`, `31:07:24.4W`, `31d 07m 24s E`
- Any input coordinate system + optional datum transformation to any output system
- Remaining Excel columns become attribute fields automatically (types inferred)
- Unparseable rows are skipped and reported with their Excel row numbers — the import never fails on one bad cell

**Use it for:** coordinates received in permits and official correspondence, client point lists, GPS/total station exports, legacy survey records.

### 05 — Attribute QA/QC Report 🆕
Scan any layer or table for data quality problems and get a formatted Excel report.

- Checks: **nulls/empties · duplicate values · whitespace problems · domain violations (coded & range) · geometry errors** (null, empty, zero area/length)
- Output: Summary sheet (counts per field per issue) + filterable Issues sheet with OID, field, value, and diagnosis
- Pairs with Tools 01/02: find issues → export → fix in Excel → apply back safely

**Use it for:** pre-migration data audits, deliverable acceptance checks, enterprise geodatabase housekeeping.

### 06 — Coordinate System Audit 🆕
Walk an entire geodatabase or folder and report the CRS of everything in it.

- Every feature class and raster: coordinate system, WKID, datum, units
- Flags **Unknown CRS** and any dataset whose **datum differs from the workspace majority**
- Excel report + datum summary

**Use it for:** inheriting legacy databases, pre-project data reviews, multi-source integration sanity checks.

### 07 — Backup Manager 🆕
Manage the timestamped backups that Tool 02 creates.

- **List** all backups with source dataset and creation time
- **Restore** — always to a *new* dataset, never overwriting anything
- **Delete** — only with an explicit confirmation checkbox

### 08 — Quick Export For Stakeholders 🆕
One click from ArcGIS Pro to the formats everyone else can open.

- **KMZ** (Google Earth) · **GeoJSON** (auto-reprojected to WGS84) · **CSV** (attributes + WGS84 coordinates)
- Each format exports independently — one failure never blocks the others

**Use it for:** management review, client deliverables, field team handoffs.

---

## Quick Start

1. Download `Bosla_Toolbox.pyt`
2. In ArcGIS Pro: **View → Catalog Pane → Toolboxes → right-click → Add Toolbox**
3. Browse to the .pyt file → **OK**
4. Run **01 – Export Attributes To Excel** on any layer — you'll have a formatted spreadsheet in seconds

## Requirements

- **ArcGIS Pro 3.x** (Windows)
- Python packages: `arcpy`, `openpyxl` — both included in the default ArcGIS Pro Python environment. No installation needed.

## Recommended Workflows

**The Excel round trip**
`01 Export → edit in Excel → 02 Apply (dry-run) → review report → 02 Apply (live)`

**Coordinate reporting for permits**
`Select boundary layer → 03 Vertices tool → pick CRS + datum transformation → DMS output → attach to report`

**Importing received coordinates**
`Paste coordinates into Excel → 04 Import → select format & CRS → verified point layer`

**Data quality cycle**
`05 QA/QC Report → fix issues via 01/02 round trip → re-run 05 to confirm clean`

## Example Applications

| Domain | Applications |
|---|---|
| **Oil & Gas** | Concession boundary reports, prospect mapping, seismic outline extraction, well planning support |
| **Geomatics / Surveying** | Survey data processing, control point verification, coordinate reporting, total station data import |
| **GIS** | Asset mapping, utility networks, land administration, enterprise geodatabase QA |
| **Engineering** | Site layout verification, infrastructure planning, route analysis |

## Screenshots

*(Add screenshots or a short GIF of the Excel round trip here — see `/docs` folder)*

## Version History

**v1.3.0** — Full geomatics suite
- 🆕 Import Points From Excel (DD/DMS/DDM/XY parsing, auto-detect, datum transformations)
- 🆕 Attribute QA/QC Report · 🆕 Coordinate System Audit · 🆕 Backup Manager · 🆕 Quick Export (KMZ/GeoJSON/CSV)

**v1.2.0** — Production hardening
- Datum transformation support + warnings in the vertices tool
- Type-safe Excel edits (key normalization, per-cell validation, no silent truncation)
- Edit-session support for enterprise geodatabases · dry-run via read-only cursor
- Metadata round-trip validation · per-feature projection (major speedup)

**v1.0.0** — Initial public release
- Export Attributes To Excel · Apply Excel Edits Back To Layer · Universal Vertices Coordinates (DD/DMS/DDM)

## Roadmap

- Excel/CSV output directly from the vertices tool
- Coordinate precision settings and lat/lon order selection
- Batch processing across multiple layers
- **ArcGIS Pro Add-In version** (ribbon UI, right-click context menu integration)

## License

MIT — free for personal, academic, and commercial use.

## About the Author

**Mohamed Abdellatief** — Geomatics & GIS Lead with 12+ years in upstream oil & gas, specializing in enterprise geospatial data architecture, full-cycle field survey operations, and AI-enabled geospatial workflows. PMP certified.

📎 [https://www.linkedin.com/in/mohamadabdellatief/]· Issues and feature requests welcome — open one on GitHub or connect on LinkedIn.

*Built with AI as a pair programmer, and 12 years of knowing exactly what to build.*

---
**Keywords:** ArcGIS Pro, ArcPy, GIS, Geomatics, Surveying, Coordinates, DMS, Excel, Python Toolbox, Spatial Analysis, Geospatial Automation, Coordinate Conversion, Datum Transformation, Data Quality, QA/QC

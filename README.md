# Geomatics Combined Tools for ArcGIS Pro

A professional ArcGIS Pro Python Toolbox designed to improve GIS, Geomatics, Surveying, and Spatial Data Management workflows.

This toolbox combines attribute editing and coordinate extraction tools into a single package and supports any coordinate system available in ArcGIS Pro through the native Coordinate System Picker. The design removes dependency on hard-coded coordinate systems, making the toolbox suitable for users worldwide.

---

# Features

## 01 - Export Attributes To Excel

Export feature class or table attributes directly to Excel (.xlsx).

### Capabilities

- Export attribute tables to Excel
- Formatted spreadsheet output
- Preserves field names
- Includes metadata sheet
- Automatically tracks key field
- Supports feature classes and standalone tables

### Use Cases

- Bulk editing in Excel
- Data review and QA/QC
- Business reporting
- Data exchange with non-GIS users

---

## 02 - Apply Excel Edits Back To Layer

Apply updates made in Excel directly back to ArcGIS Pro.

### Capabilities

- Key field matching
- Update selected fields
- Dry-run mode
- Automatic backup creation
- Null handling options
- Bulk attribute updates

### Use Cases

- Mass attribute corrections
- Database maintenance
- QA/QC workflows
- Vendor data updates

---

## 03 - Universal Vertices Coordinates Tool

Convert polygon and polyline vertices into point features while exporting coordinates in multiple formats.

### Capabilities

- Supports Polygon features
- Supports Polyline features
- Supports Multipart features
- Creates vertex point feature class
- Exports projected coordinates
- Exports geographic coordinates
- Supports Z values
- Supports custom coordinate systems
- Supports any ArcGIS Pro spatial reference

### Coordinate Formats

#### Decimal Degrees

Example:

31.123456

29.987654

#### Degrees Minutes Seconds (DMS)

Example:

31° 07' 24.441" E

29° 59' 12.645" N

#### Degrees Decimal Minutes (DDM)

Example:

31° 07.40735' E

29° 59.21071' N

---

# Universal Coordinate System Support

This version does not contain any predefined coordinate systems.

Users can select any coordinate system directly from the ArcGIS Pro Coordinate System Picker.

### Supported Examples

- WGS 1984
- UTM Zones
- Egypt 1907
- State Plane
- Lambert Conformal Conic
- British National Grid
- National Grids
- Local Coordinate Systems
- Engineering Coordinate Systems
- Custom PRJ Files
- Project Favorites

No code modifications are required to support additional coordinate systems.

---

# Installation

## Method 1 - Add Toolbox to ArcGIS Pro

1. Download:

   GeomaticsCombinedTools_UniversalCS.pyt

2. Open ArcGIS Pro

3. Open:

   View → Catalog Pane

4. Right-click:

   Toolboxes

5. Select:

   Add Toolbox

6. Browse to:

   GeomaticsCombinedTools_UniversalCS.pyt

7. Click OK

The toolbox will appear under:

Geomatics Combined Tools - Universal CS

---

# Requirements

## Software

- ArcGIS Pro 3.x

## Python Packages

Included by default in most ArcGIS Pro installations:

- arcpy
- openpyxl

---

# Recommended Workflow

## Attribute Editing

Step 1:

Export Attributes To Excel

<img width="576" height="372" alt="1" src="https://github.com/user-attachments/assets/36eb5d95-d5a3-48a7-b650-fcf15efdffdf" />


Step 2:

Edit attributes in Excel



Step 3:

Save Excel file

↓

Step 4:

Apply Excel Edits Back To Layer

<img width="563" height="432" alt="2" src="https://github.com/user-attachments/assets/c1145c5e-6c69-4655-a96b-e2a6c0ae50a1" />

---

## Coordinate Extraction

Step 1:

Select Polygon or Polyline layer

<img width="575" height="514" alt="3" src="https://github.com/user-attachments/assets/bba9a6f5-f2b6-42f8-9107-d14b7223a53d" />


Step 2:

Run:

Outline To Vertices Coordinates – Universal CS

↓

Step 3:

Select desired:

- Projected Coordinate System
- Geographic Coordinate System

↓

Step 4:

Choose output format:

- Decimal Degrees
- DMS
- DDM

↓

Step 5:

Create output vertices feature class

---

# Example Applications

## GIS

- Asset Mapping
- Utility Networks
- Land Administration
- Environmental Studies

## Geomatics

- Survey Data Processing
- Control Point Verification
- Coordinate Reporting

## Oil & Gas

- Prospect Mapping
- Seismic Outline Extraction
- Concession Boundary Analysis
- Well Planning Support

## Engineering

- Site Layout Verification
- Infrastructure Planning
- Route Analysis

---

# Screenshots

## Export Attributes To Excel

![Export Attributes](images/ExportExcel.png)

---

## Apply Excel Edits Back To Layer

![Apply Edits](images/ApplyExcelEdits.png)

---

## Universal Coordinate System Picker

![Coordinate Picker](images/UniversalCoordinatePicker.png)

---

## Vertices Coordinate Output

![Vertices Output](images/VerticesOutput.png)

---

# Version History

## v1.0.0

Initial Public Release

Features:

- Export Attributes To Excel
- Apply Excel Edits Back To Layer
- Universal Vertices Coordinates Tool
- Decimal Degrees Export
- DMS Export
- DDM Export
- Universal Coordinate System Picker
- Spatial Reference Independence

---

# Future Roadmap

Planned Enhancements

- Excel Output from Vertices Tool
- CSV Output Support
- Coordinate Precision Settings
- Latitude/Longitude Order Selection
- Coordinate Labels
- Attribute Transfer to Vertices
- Batch Processing
- ArcGIS Pro Add-In Version
- Right-Click Layer Context Menu Integration

---

# License

MIT License

---

# Author

Mohamed Abdellatief

Geomatics / GIS Specialist

ArcGIS Pro • Python • Surveying • Spatial Data Management

---

# Keywords

ArcGIS Pro, ArcPy, GIS, Geomatics, Surveying, Coordinates, Excel, Python Toolbox, Spatial Analysis, Geospatial Automation, Coordinate Conversion, Data Management

# -*- coding: utf-8 -*-
"""
Geomatics Combined Tools for ArcGIS Pro
Version: 1.3.0 (Universal Coordinate System edition - full geomatics suite)

Tools included:
  1. Export Attributes To Excel
  2. Apply Excel Edits Back To Layer
  3. Outline To Vertices Coordinates - Universal Coordinate System Picker
  4. Import Points From Excel (DD / DMS / DDM / Projected XY parsing)
  5. Attribute QA/QC Report
  6. Coordinate System Audit
  7. Backup Manager (list / restore / delete Tool 2 backups)
  8. Quick Export For Stakeholders (KMZ / GeoJSON / CSV)

Changes in 1.3.0:
  - NEW Tool 4 - Import Points From Excel: the reverse direction of the
    round trip. Parses coordinate text in Decimal Degrees, DMS
    (e.g. 31d 07' 24.441" E), DDM, or projected XY, in any coordinate
    system, with optional datum transformation and automatic attribute
    field creation from the remaining Excel columns.
  - NEW Tool 5 - Attribute QA/QC Report: scans a layer/table for nulls,
    duplicate values, whitespace problems, domain violations, and geometry
    problems, and writes a formatted Excel report (Summary + Issues).
  - NEW Tool 6 - Coordinate System Audit: walks a workspace and reports the
    coordinate system, WKID, datum, and units of every feature class and
    raster, flagging Unknown CRS and datum mismatches.
  - NEW Tool 7 - Backup Manager: lists, restores, and deletes the
    timestamped backups created by Tool 2.
  - NEW Tool 8 - Quick Export For Stakeholders: one-click export to
    KMZ, GeoJSON (WGS84), and/or CSV for non-GIS colleagues.

Changes in 1.2.0 (code review hardening):
  Tool 2 - Apply Excel Edits Back To Layer:
    - Key values from Excel are normalized to the key field's data type
      (openpyxl often returns integers as floats, e.g. 1.0, which previously
      caused silent "Matched rows: 0" results).
    - All incoming Excel values are cast to the target field's data type
      before comparison and update. Invalid values (wrong type, text longer
      than the field length, unparseable dates) are skipped and reported
      instead of failing mid-update.
    - Updates now run inside an edit session (arcpy.da.Editor) when possible,
      enabling versioned enterprise geodatabase data and providing rollback
      on failure. Falls back with a warning where an edit session cannot
      be opened.
    - Dry-run mode now uses a read-only SearchCursor (no schema locks,
      works on data the user cannot edit).
    - The tool now prefers the worksheet named "Attributes", then the first
      visible sheet (previously: blindly the first sheet).
    - The hidden _GIS_METADATA sheet written by Tool 1 is now read back:
      the tool warns if the Excel file appears to come from a different
      source dataset or was exported with a different key field.
  Tool 3 - Outline To Vertices Coordinates:
    - NEW: optional geographic (datum) transformation parameters for both
      the projected and geographic outputs. The pick lists are populated
      automatically from arcpy.ListTransformations() based on the input
      layer and the selected target coordinate systems. A warning is raised
      when the datums differ and no transformation is selected.
      (Without this, e.g. Egypt 1907 <-> WGS 1984 conversions could be
      offset by tens of meters.)
    - Performance: geometries are now projected once per feature instead of
      once per vertex (roughly O(features) instead of O(vertices) projection
      calls - a large speedup on dense boundaries).
    - Output points are written via SHAPE@XY (no per-vertex geometry
      construction).
  Helpers:
    - _make_backup: shapefile/folder workspaces are now backed up into the
      scratch geodatabase to avoid file-extension issues; the backup path
      is always reported.

Prepared for: Mohamed Abdellatief
License: MIT
"""

import arcpy
import os
import csv
import math
import datetime

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except Exception:
    openpyxl = None


__version__ = "1.3.0"

METADATA_SHEET = "_GIS_METADATA"
ATTRIBUTES_SHEET = "Attributes"

INTEGER_FIELD_TYPES = ("Integer", "SmallInteger", "BigInteger", "OID")
FLOAT_FIELD_TYPES = ("Double", "Single")


class Toolbox(object):
    def __init__(self):
        self.label = "Geomatics Combined Tools - Universal CS"
        self.alias = "geomatics_combined_universal_cs"
        self.tools = [
            ExportAttributesToExcel,
            ApplyExcelEditsBackToLayer,
            OutlineToVerticesCoordinatesUniversalCS,
            ImportPointsFromExcel,
            AttributeQAQCReport,
            CoordinateSystemAudit,
            BackupManager,
            QuickExportForStakeholders
        ]


# -----------------------------------------------------------------------------
# Common helper functions
# -----------------------------------------------------------------------------

def _require_openpyxl():
    if openpyxl is None:
        raise arcpy.ExecuteError(
            "openpyxl is required to read/write XLSX files. "
            "It is usually included in the ArcGIS Pro Python environment."
        )


def _safe_excel_value(value):
    if value is None:
        return None
    if isinstance(value, (str, int, float, bool, datetime.datetime, datetime.date)):
        return value
    return str(value)


def _parse_multivalue(value_as_text, default_values=None):
    if default_values is None:
        default_values = []
    if not value_as_text:
        return default_values
    return [v.strip().strip("'").strip('"') for v in value_as_text.split(";") if v.strip()]


def _editable_update_fields(dataset, key_field, selected_fields=None):
    selected_upper = set([x.upper() for x in selected_fields]) if selected_fields else None
    editable = []

    for field in arcpy.ListFields(dataset):
        if field.name == key_field:
            continue
        if field.type in ["OID", "Geometry", "GlobalID", "GUID", "Blob", "Raster"]:
            continue
        if not field.editable:
            continue
        if field.required:
            continue
        if selected_upper and field.name.upper() not in selected_upper:
            continue
        editable.append(field.name)

    return editable


def _make_backup(dataset):
    """Copy the dataset to a timestamped backup and return its path.

    File-system workspaces (shapefiles, standalone dBASE tables) are backed
    up into the scratch geodatabase to avoid file-extension pitfalls.
    """
    desc = arcpy.Describe(dataset)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = os.path.splitext(os.path.basename(desc.name))[0]

    workspace = getattr(desc, "path", None) or arcpy.env.scratchGDB
    try:
        workspace_desc = arcpy.Describe(workspace)
        if getattr(workspace_desc, "dataType", "") == "FeatureDataset":
            workspace = getattr(workspace_desc, "path", workspace)
            workspace_desc = arcpy.Describe(workspace)
        if getattr(workspace_desc, "workspaceType", "") == "FileSystem":
            workspace = arcpy.env.scratchGDB
    except Exception:
        workspace = arcpy.env.scratchGDB

    backup_name = arcpy.ValidateTableName("{}_backup_{}".format(base_name, stamp), workspace)
    backup_path = os.path.join(workspace, backup_name)

    if hasattr(desc, "shapeType"):
        arcpy.management.CopyFeatures(dataset, backup_path)
    else:
        arcpy.management.CopyRows(dataset, backup_path)

    return backup_path


def _workspace_for_editor(dataset):
    """Return the workspace path suitable for arcpy.da.Editor, or None."""
    try:
        desc = arcpy.Describe(dataset)
        path = getattr(desc, "path", None)
        if not path:
            return None
        parent_desc = arcpy.Describe(path)
        if getattr(parent_desc, "dataType", "") == "FeatureDataset":
            path = getattr(parent_desc, "path", path)
        return path
    except Exception:
        return None


def _normalize_key(value, key_field_type):
    """Normalize a key value (from Excel or the database) so both sides match.

    openpyxl frequently returns whole numbers as floats (1.0). For integer
    key fields, both Excel and database keys are normalized to int. String
    keys are stripped of surrounding whitespace.
    """
    if value is None:
        return None
    if key_field_type in INTEGER_FIELD_TYPES:
        try:
            as_float = float(value)
            if abs(as_float - round(as_float)) < 1e-9:
                return int(round(as_float))
        except (TypeError, ValueError):
            pass
        return value
    if key_field_type == "String":
        return str(value).strip()
    return value


def _cast_value_for_field(value, field):
    """Cast an Excel cell value to the target field's data type.

    Raises ValueError with a human-readable reason when the value cannot be
    represented safely (e.g. non-integer number for an integer field, text
    longer than the field length, unparseable date). Silent truncation or
    rounding is deliberately avoided - survey data should never be altered
    implicitly.
    """
    if value is None:
        return None

    field_type = field.type

    if field_type in INTEGER_FIELD_TYPES:
        if isinstance(value, bool):
            raise ValueError("boolean value not valid for integer field")
        as_float = float(value)
        if abs(as_float - round(as_float)) > 1e-9:
            raise ValueError("value {} is not a whole number".format(value))
        return int(round(as_float))

    if field_type in FLOAT_FIELD_TYPES:
        return float(value)

    if field_type == "String":
        text = value if isinstance(value, str) else str(value)
        if field.length and len(text) > field.length:
            raise ValueError(
                "text length {} exceeds field length {}".format(len(text), field.length)
            )
        return text

    if field_type == "Date":
        if isinstance(value, datetime.datetime):
            return value
        if isinstance(value, datetime.date):
            return datetime.datetime(value.year, value.month, value.day)
        if isinstance(value, str):
            candidate = value.strip()
            for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d",
                        "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%m/%d/%Y"):
                try:
                    return datetime.datetime.strptime(candidate, fmt)
                except ValueError:
                    continue
            raise ValueError("unrecognized date format: {}".format(value))
        raise ValueError("cannot convert {} to a date".format(type(value).__name__))

    return value


def _gcs_name(spatial_reference):
    """Return the underlying geographic coordinate system name of an SR."""
    if spatial_reference is None:
        return None
    try:
        gcs = spatial_reference.GCS
        if gcs is not None and gcs.name:
            return gcs.name
    except Exception:
        pass
    try:
        return spatial_reference.GCSName or spatial_reference.name
    except Exception:
        return spatial_reference.name


# -----------------------------------------------------------------------------
# Tool 1: Export Attributes to Excel
# -----------------------------------------------------------------------------
class ExportAttributesToExcel(object):
    def __init__(self):
        self.label = "01 - Export Attributes To Excel"
        self.description = "Export a layer/table attribute table to editable XLSX."
        self.canRunInBackground = False

    def getParameterInfo(self):
        in_table = arcpy.Parameter(
            displayName="Input Layer or Table",
            name="in_table",
            datatype="GPTableView",
            parameterType="Required",
            direction="Input"
        )

        out_excel = arcpy.Parameter(
            displayName="Output Excel File (.xlsx)",
            name="out_excel",
            datatype="DEFile",
            parameterType="Required",
            direction="Output"
        )
        out_excel.filter.list = ["xlsx"]

        key_field = arcpy.Parameter(
            displayName="Key Field for Re-import",
            name="key_field",
            datatype="Field",
            parameterType="Optional",
            direction="Input"
        )
        key_field.parameterDependencies = [in_table.name]

        open_after_export = arcpy.Parameter(
            displayName="Open Excel After Export",
            name="open_after_export",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        open_after_export.value = False

        return [in_table, out_excel, key_field, open_after_export]

    def updateParameters(self, parameters):
        if parameters[0].valueAsText and not parameters[2].altered:
            try:
                desc = arcpy.Describe(parameters[0].valueAsText)
                if hasattr(desc, "OIDFieldName"):
                    parameters[2].value = desc.OIDFieldName
            except Exception:
                pass
        return

    def execute(self, parameters, messages):
        _require_openpyxl()

        in_table = parameters[0].valueAsText
        out_excel = parameters[1].valueAsText
        key_field = parameters[2].valueAsText
        open_after_export = bool(parameters[3].value)

        if not out_excel.lower().endswith(".xlsx"):
            out_excel += ".xlsx"

        desc = arcpy.Describe(in_table)

        if not key_field:
            key_field = getattr(desc, "OIDFieldName", None)
        if not key_field:
            raise arcpy.ExecuteError("Please choose a key field for re-import.")

        field_names = []
        for field in arcpy.ListFields(in_table):
            if field.type not in ["Geometry", "Blob", "Raster"]:
                field_names.append(field.name)

        if key_field in field_names:
            field_names.remove(key_field)
        field_names.insert(0, key_field)

        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = ATTRIBUTES_SHEET

        for col_index, field_name in enumerate(field_names, start=1):
            cell = worksheet.cell(row=1, column=col_index, value=field_name)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center")

        row_count = 0
        with arcpy.da.SearchCursor(in_table, field_names) as cursor:
            for row_index, row in enumerate(cursor, start=2):
                for col_index, value in enumerate(row, start=1):
                    worksheet.cell(row=row_index, column=col_index, value=_safe_excel_value(value))
                row_count += 1

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        for col_index, field_name in enumerate(field_names, start=1):
            worksheet.column_dimensions[get_column_letter(col_index)].width = min(45, max(12, len(field_name) + 4))

        metadata = workbook.create_sheet(METADATA_SHEET)
        metadata.sheet_state = "hidden"
        metadata["A1"] = "source"
        metadata["B1"] = getattr(desc, "catalogPath", in_table)
        metadata["A2"] = "key_field"
        metadata["B2"] = key_field
        metadata["A3"] = "instruction"
        metadata["B3"] = "Do not change key values. Edit attributes only. Save and close Excel before applying back."
        metadata["A4"] = "toolbox_version"
        metadata["B4"] = __version__

        folder = os.path.dirname(out_excel)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)

        workbook.save(out_excel)
        workbook.close()

        arcpy.AddMessage("Export completed successfully.")
        arcpy.AddMessage("Rows exported: {}".format(row_count))
        arcpy.AddMessage("Key field: {}".format(key_field))
        arcpy.AddMessage("Excel file: {}".format(out_excel))

        if open_after_export:
            try:
                os.startfile(out_excel)
            except Exception:
                arcpy.AddWarning("Excel file was created but could not be opened automatically.")

        return


# -----------------------------------------------------------------------------
# Tool 2: Apply Excel Edits Back to Layer
# -----------------------------------------------------------------------------
class ApplyExcelEditsBackToLayer(object):
    def __init__(self):
        self.label = "02 - Apply Excel Edits Back To Layer"
        self.description = (
            "Apply edited XLSX values back to a layer/table using a key field. "
            "Values are validated and cast to the target field types; updates "
            "run inside an edit session where possible. Dry-run is enabled by "
            "default."
        )
        self.canRunInBackground = False

    def getParameterInfo(self):
        target_table = arcpy.Parameter(
            displayName="Target Layer or Table",
            name="target_table",
            datatype="GPTableView",
            parameterType="Required",
            direction="Input"
        )

        in_excel = arcpy.Parameter(
            displayName="Edited Excel File (.xlsx)",
            name="in_excel",
            datatype="DEFile",
            parameterType="Required",
            direction="Input"
        )
        in_excel.filter.list = ["xlsx"]

        key_field = arcpy.Parameter(
            displayName="Key Field",
            name="key_field",
            datatype="Field",
            parameterType="Required",
            direction="Input"
        )
        key_field.parameterDependencies = [target_table.name]

        fields_to_update = arcpy.Parameter(
            displayName="Fields to Update (blank = all editable matching fields)",
            name="fields_to_update",
            datatype="Field",
            parameterType="Optional",
            direction="Input",
            multiValue=True
        )
        fields_to_update.parameterDependencies = [target_table.name]

        treat_blanks_as_null = arcpy.Parameter(
            displayName="Treat Blank Excel Cells as Null",
            name="treat_blanks_as_null",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        treat_blanks_as_null.value = False

        create_backup = arcpy.Parameter(
            displayName="Create Backup Before Update",
            name="create_backup",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        create_backup.value = True

        dry_run = arcpy.Parameter(
            displayName="Dry Run Only - Report Changes Without Updating",
            name="dry_run",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        dry_run.value = True

        return [target_table, in_excel, key_field, fields_to_update, treat_blanks_as_null, create_backup, dry_run]

    def updateParameters(self, parameters):
        if parameters[0].valueAsText and not parameters[2].altered:
            try:
                desc = arcpy.Describe(parameters[0].valueAsText)
                if hasattr(desc, "OIDFieldName"):
                    parameters[2].value = desc.OIDFieldName
            except Exception:
                pass
        return

    # -- internal helpers ------------------------------------------------

    def _pick_worksheet(self, workbook):
        """Prefer the 'Attributes' sheet, then the first visible sheet."""
        if ATTRIBUTES_SHEET in workbook.sheetnames:
            return workbook[ATTRIBUTES_SHEET]
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            if sheet.sheet_state == "visible":
                return sheet
        return workbook[workbook.sheetnames[0]]

    def _read_metadata(self, workbook):
        metadata = {}
        if METADATA_SHEET in workbook.sheetnames:
            sheet = workbook[METADATA_SHEET]
            for row in sheet.iter_rows(min_row=1, max_col=2, values_only=True):
                if row and row[0] is not None:
                    metadata[str(row[0]).strip()] = row[1]
        return metadata

    def _validate_metadata(self, metadata, target_table, key_field):
        """Warn when the Excel file appears to come from another dataset."""
        if not metadata:
            return
        meta_key = metadata.get("key_field")
        if meta_key and str(meta_key).upper() != key_field.upper():
            arcpy.AddWarning(
                "This Excel file was exported with key field '{}' but you selected '{}'. "
                "Row matching may be incorrect.".format(meta_key, key_field)
            )
        meta_source = metadata.get("source")
        if meta_source:
            try:
                target_name = os.path.basename(arcpy.Describe(target_table).catalogPath)
                source_name = os.path.basename(str(meta_source))
                if target_name.lower() != source_name.lower():
                    arcpy.AddWarning(
                        "This Excel file was exported from '{}' but the target is '{}'. "
                        "Verify you selected the intended layer.".format(source_name, target_name)
                    )
            except Exception:
                pass

    def _apply(self, cursor_type, target_table, cursor_fields, excel_edits,
               update_fields, field_lookup, header_upper, key_field_type,
               treat_blanks_as_null, dry_run):
        """Run the compare/update loop with the given cursor type.

        Returns (matched_rows, changed_cells, skipped_blank_cells, cast_errors).
        cast_errors is a list of (key, field, value, reason) tuples.
        """
        matched_rows = 0
        changed_cells = 0
        skipped_blank_cells = 0
        cast_errors = []

        with cursor_type(target_table, cursor_fields) as cursor:
            for db_row in cursor:
                db_key = _normalize_key(db_row[0], key_field_type)
                if db_key not in excel_edits:
                    continue

                matched_rows += 1
                excel_row = excel_edits[db_key]
                new_row = list(db_row)
                row_changed = False

                for idx, field_name in enumerate(update_fields, start=1):
                    excel_header_name = header_upper.get(field_name.upper())
                    raw_value = excel_row.get(excel_header_name)

                    if raw_value in [None, ""]:
                        if not treat_blanks_as_null:
                            skipped_blank_cells += 1
                            continue
                        new_value = None
                    else:
                        try:
                            new_value = _cast_value_for_field(raw_value, field_lookup[field_name.upper()])
                        except (ValueError, TypeError) as cast_error:
                            cast_errors.append((db_key, field_name, raw_value, str(cast_error)))
                            continue

                    if db_row[idx] != new_value:
                        new_row[idx] = new_value
                        row_changed = True
                        changed_cells += 1

                if row_changed and not dry_run:
                    cursor.updateRow(new_row)

        return matched_rows, changed_cells, skipped_blank_cells, cast_errors

    # -- execute -----------------------------------------------------------

    def execute(self, parameters, messages):
        _require_openpyxl()

        target_table = parameters[0].valueAsText
        in_excel = parameters[1].valueAsText
        key_field = parameters[2].valueAsText
        selected_fields = _parse_multivalue(parameters[3].valueAsText, [])
        treat_blanks_as_null = bool(parameters[4].value)
        create_backup = bool(parameters[5].value)
        dry_run = bool(parameters[6].value)

        if not os.path.exists(in_excel):
            raise arcpy.ExecuteError("Excel file does not exist: {}".format(in_excel))

        field_lookup = {f.name.upper(): f for f in arcpy.ListFields(target_table)}
        if key_field.upper() not in field_lookup:
            raise arcpy.ExecuteError("Key field not found in target: {}".format(key_field))
        key_field_type = field_lookup[key_field.upper()].type

        workbook = openpyxl.load_workbook(in_excel, data_only=True)
        worksheet = self._pick_worksheet(workbook)
        self._validate_metadata(self._read_metadata(workbook), target_table, key_field)
        arcpy.AddMessage("Reading worksheet: {}".format(worksheet.title))

        header = [str(cell.value).strip() if cell.value is not None else "" for cell in worksheet[1]]
        header_upper = {field.upper(): field for field in header if field}

        if key_field.upper() not in header_upper:
            workbook.close()
            raise arcpy.ExecuteError("Excel file does not contain key field: {}".format(key_field))

        excel_key_header = header_upper[key_field.upper()]
        excel_edits = {}
        duplicate_keys = set()

        for row_index in range(2, worksheet.max_row + 1):
            record = {}
            has_value = False
            for col_index, field_name in enumerate(header, start=1):
                if not field_name:
                    continue
                value = worksheet.cell(row_index, col_index).value
                if value is not None:
                    has_value = True
                record[field_name] = value

            if has_value:
                key_value = _normalize_key(record.get(excel_key_header), key_field_type)
                if key_value in [None, ""]:
                    continue
                if key_value in excel_edits:
                    duplicate_keys.add(key_value)
                excel_edits[key_value] = record

        workbook.close()

        if duplicate_keys:
            raise arcpy.ExecuteError(
                "Duplicate key value(s) found in Excel: {}".format(
                    ", ".join(str(k) for k in list(duplicate_keys)[:5])
                )
            )

        update_fields = []
        for field_name in _editable_update_fields(target_table, key_field, selected_fields):
            if field_name.upper() in header_upper:
                update_fields.append(field_name)

        if not update_fields:
            raise arcpy.ExecuteError("No editable matching fields were found between Excel and the target table.")

        if create_backup and not dry_run:
            backup_path = _make_backup(target_table)
            arcpy.AddMessage("Backup created: {}".format(backup_path))

        cursor_fields = [key_field] + update_fields

        if dry_run:
            results = self._apply(
                arcpy.da.SearchCursor, target_table, cursor_fields, excel_edits,
                update_fields, field_lookup, header_upper, key_field_type,
                treat_blanks_as_null, dry_run=True
            )
        else:
            workspace = _workspace_for_editor(target_table)
            results = None
            if workspace:
                try:
                    with arcpy.da.Editor(workspace):
                        results = self._apply(
                            arcpy.da.UpdateCursor, target_table, cursor_fields, excel_edits,
                            update_fields, field_lookup, header_upper, key_field_type,
                            treat_blanks_as_null, dry_run=False
                        )
                    arcpy.AddMessage("Updates applied inside an edit session (workspace: {}).".format(workspace))
                except (RuntimeError, arcpy.ExecuteError) as edit_error:
                    arcpy.AddWarning(
                        "Could not open an edit session ({}). "
                        "Falling back to a direct update - the automatic backup is your rollback.".format(edit_error)
                    )
            if results is None:
                results = self._apply(
                    arcpy.da.UpdateCursor, target_table, cursor_fields, excel_edits,
                    update_fields, field_lookup, header_upper, key_field_type,
                    treat_blanks_as_null, dry_run=False
                )

        matched_rows, changed_cells, skipped_blank_cells, cast_errors = results

        arcpy.AddMessage("Excel apply process completed.")
        arcpy.AddMessage("Excel rows read: {}".format(len(excel_edits)))
        arcpy.AddMessage("Matched rows: {}".format(matched_rows))
        arcpy.AddMessage("Changed cells detected: {}".format(changed_cells))
        arcpy.AddMessage("Blank cells skipped: {}".format(skipped_blank_cells))
        arcpy.AddMessage("Updated fields: {}".format(", ".join(update_fields)))

        if matched_rows == 0 and excel_edits:
            arcpy.AddWarning(
                "No rows matched. Check that the key field '{}' contains the same values "
                "in Excel and in the target table.".format(key_field)
            )

        if cast_errors:
            arcpy.AddWarning("{} cell(s) had invalid values and were skipped:".format(len(cast_errors)))
            for key, field_name, value, reason in cast_errors[:20]:
                arcpy.AddWarning("  Key {} | field {} | value '{}' | {}".format(key, field_name, value, reason))
            if len(cast_errors) > 20:
                arcpy.AddWarning("  ... and {} more.".format(len(cast_errors) - 20))

        if dry_run:
            arcpy.AddWarning("Dry run only: no edits were applied. Uncheck 'Dry Run' to apply the changes above.")
        else:
            arcpy.AddMessage("Edits applied to target layer/table.")

        return


# -----------------------------------------------------------------------------
# Tool 3: Universal vertices coordinate tool
# -----------------------------------------------------------------------------
class OutlineToVerticesCoordinatesUniversalCS(object):
    def __init__(self):
        self.label = "03 - Outline To Vertices Coordinates - Universal CS"
        self.description = (
            "Convert polyline/polygon outline vertices to points. "
            "No preferred coordinate systems are hard-coded. Users select any projected "
            "and/or geographic coordinate system from ArcGIS Pro's coordinate system picker, "
            "with optional datum transformations for survey-grade accuracy."
        )
        self.canRunInBackground = False

    def geographic_format_options(self):
        return [
            "Decimal Degrees",
            "Degrees Minutes Seconds (DMS)",
            "Degrees Decimal Minutes (DDM)"
        ]

    def _hemisphere(self, value, coord_type):
        if coord_type == "LAT":
            return "N" if value >= 0 else "S"
        return "E" if value >= 0 else "W"

    def _decimal_degrees_to_dms(self, value, coord_type):
        hemi = self._hemisphere(value, coord_type)
        value_abs = abs(float(value))
        degrees = int(math.floor(value_abs))
        minutes_float = (value_abs - degrees) * 60.0
        minutes = int(math.floor(minutes_float))
        seconds = round((minutes_float - minutes) * 60.0, 3)

        if seconds >= 60.0:
            seconds = 0.0
            minutes += 1
        if minutes >= 60:
            minutes = 0
            degrees += 1

        return "{}{} {:02d}' {:06.3f}\" {}".format(degrees, chr(176), minutes, seconds, hemi)

    def _decimal_degrees_to_ddm(self, value, coord_type):
        hemi = self._hemisphere(value, coord_type)
        value_abs = abs(float(value))
        degrees = int(math.floor(value_abs))
        decimal_minutes = round((value_abs - degrees) * 60.0, 5)

        if decimal_minutes >= 60.0:
            decimal_minutes = 0.0
            degrees += 1

        return "{}{} {:08.5f}' {}".format(degrees, chr(176), decimal_minutes, hemi)

    def _safe_factory_code(self, spatial_reference):
        try:
            code = spatial_reference.factoryCode
            return code if code else 0
        except Exception:
            return 0

    def getParameterInfo(self):
        in_features = arcpy.Parameter(
            displayName="Input Line or Polygon Features",
            name="in_features",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input"
        )
        in_features.filter.list = ["Polyline", "Polygon"]

        out_points = arcpy.Parameter(
            displayName="Output Vertices Point Feature Class",
            name="out_points",
            datatype="DEFeatureClass",
            parameterType="Required",
            direction="Output"
        )

        add_projected = arcpy.Parameter(
            displayName="Add Projected Coordinates",
            name="add_projected",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        add_projected.value = True

        projected_sr = arcpy.Parameter(
            displayName="Projected Coordinate System - User Selected",
            name="projected_sr",
            datatype="GPSpatialReference",
            parameterType="Optional",
            direction="Input"
        )

        projected_transformation = arcpy.Parameter(
            displayName="Datum Transformation for Projected Output (recommended when datums differ)",
            name="projected_transformation",
            datatype="GPString",
            parameterType="Optional",
            direction="Input"
        )
        projected_transformation.filter.type = "ValueList"
        projected_transformation.filter.list = []

        add_geographic = arcpy.Parameter(
            displayName="Add Geographic Coordinates",
            name="add_geographic",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        add_geographic.value = True

        geographic_sr = arcpy.Parameter(
            displayName="Geographic Coordinate System - User Selected",
            name="geographic_sr",
            datatype="GPSpatialReference",
            parameterType="Optional",
            direction="Input"
        )

        geographic_transformation = arcpy.Parameter(
            displayName="Datum Transformation for Geographic Output (recommended when datums differ)",
            name="geographic_transformation",
            datatype="GPString",
            parameterType="Optional",
            direction="Input"
        )
        geographic_transformation.filter.type = "ValueList"
        geographic_transformation.filter.list = []

        geographic_formats = arcpy.Parameter(
            displayName="Geographic Coordinate Number Format(s)",
            name="geographic_formats",
            datatype="GPString",
            parameterType="Optional",
            direction="Input",
            multiValue=True
        )
        geographic_formats.filter.type = "ValueList"
        geographic_formats.filter.list = self.geographic_format_options()
        geographic_formats.values = [["Decimal Degrees"], ["Degrees Minutes Seconds (DMS)"], ["Degrees Decimal Minutes (DDM)"]]

        keep_closing_vertex = arcpy.Parameter(
            displayName="Keep Duplicate Polygon Closing Vertex",
            name="keep_closing_vertex",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        keep_closing_vertex.value = False

        include_z = arcpy.Parameter(
            displayName="Include Z Value If Available",
            name="include_z",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        include_z.value = False

        output_geometry_sr_option = arcpy.Parameter(
            displayName="Output Point Geometry Coordinate System",
            name="output_geometry_sr_option",
            datatype="GPString",
            parameterType="Optional",
            direction="Input"
        )
        output_geometry_sr_option.filter.type = "ValueList"
        output_geometry_sr_option.filter.list = [
            "Same as Input Layer",
            "Use Selected Projected Coordinate System",
            "Use Selected Geographic Coordinate System"
        ]
        output_geometry_sr_option.value = "Same as Input Layer"

        return [
            in_features,                 # 0
            out_points,                  # 1
            add_projected,               # 2
            projected_sr,                # 3
            projected_transformation,    # 4
            add_geographic,              # 5
            geographic_sr,               # 6
            geographic_transformation,   # 7
            geographic_formats,          # 8
            keep_closing_vertex,         # 9
            include_z,                   # 10
            output_geometry_sr_option    # 11
        ]

    def _input_spatial_reference(self, parameters):
        try:
            if parameters[0].value:
                return arcpy.Describe(parameters[0].value).spatialReference
        except Exception:
            pass
        return None

    def updateParameters(self, parameters):
        parameters[3].enabled = bool(parameters[2].value)
        parameters[4].enabled = bool(parameters[2].value)
        parameters[6].enabled = bool(parameters[5].value)
        parameters[7].enabled = bool(parameters[5].value)
        parameters[8].enabled = bool(parameters[5].value)

        input_sr = self._input_spatial_reference(parameters)
        if input_sr and input_sr.name != "Unknown":
            try:
                if parameters[3].value:
                    transformations = arcpy.ListTransformations(input_sr, parameters[3].value) or []
                    parameters[4].filter.list = transformations
            except Exception:
                pass
            try:
                if parameters[6].value:
                    transformations = arcpy.ListTransformations(input_sr, parameters[6].value) or []
                    parameters[7].filter.list = transformations
            except Exception:
                pass
        return

    def updateMessages(self, parameters):
        add_projected = bool(parameters[2].value)
        add_geographic = bool(parameters[5].value)

        if not add_projected and not add_geographic:
            parameters[2].setErrorMessage("Enable at least one of Projected or Geographic coordinates.")
            parameters[5].setErrorMessage("Enable at least one of Projected or Geographic coordinates.")
            return

        if add_projected and parameters[3].value is not None:
            try:
                if parameters[3].value.type != "Projected":
                    parameters[3].setErrorMessage("Please select a PROJECTED coordinate system.")
            except Exception:
                pass

        if add_geographic and parameters[6].value is not None:
            try:
                if parameters[6].value.type != "Geographic":
                    parameters[6].setErrorMessage("Please select a GEOGRAPHIC coordinate system.")
            except Exception:
                pass

        input_sr = self._input_spatial_reference(parameters)
        if input_sr and input_sr.name != "Unknown":
            input_gcs = _gcs_name(input_sr)

            if add_projected and parameters[3].value is not None and not parameters[4].valueAsText:
                if input_gcs and _gcs_name(parameters[3].value) and input_gcs != _gcs_name(parameters[3].value):
                    parameters[4].setWarningMessage(
                        "The input datum ({}) differs from the target datum ({}). "
                        "Select a datum transformation, or coordinates may be offset "
                        "by tens of meters or more.".format(input_gcs, _gcs_name(parameters[3].value))
                    )

            if add_geographic and parameters[6].value is not None and not parameters[7].valueAsText:
                if input_gcs and _gcs_name(parameters[6].value) and input_gcs != _gcs_name(parameters[6].value):
                    parameters[7].setWarningMessage(
                        "The input datum ({}) differs from the target datum ({}). "
                        "Select a datum transformation, or coordinates may be offset "
                        "by tens of meters or more.".format(input_gcs, _gcs_name(parameters[6].value))
                    )
        return

    @staticmethod
    def _part_points(geometry, part_index):
        """Return the non-null points of one geometry part as a list."""
        return [point for point in geometry[part_index] if point is not None]

    def execute(self, parameters, messages):
        arcpy.env.overwriteOutput = True

        in_features = parameters[0].valueAsText
        out_points = parameters[1].valueAsText
        add_projected = bool(parameters[2].value)
        projected_sr = parameters[3].value
        projected_transformation = (parameters[4].valueAsText or "").strip() or None
        add_geographic = bool(parameters[5].value)
        geographic_sr = parameters[6].value
        geographic_transformation = (parameters[7].valueAsText or "").strip() or None
        geographic_formats = _parse_multivalue(parameters[8].valueAsText, ["Decimal Degrees"])
        keep_closing_vertex = bool(parameters[9].value)
        include_z = bool(parameters[10].value)
        output_geometry_sr_option = parameters[11].valueAsText

        desc = arcpy.Describe(in_features)
        shape_type = desc.shapeType
        input_sr = desc.spatialReference

        if shape_type not in ["Polyline", "Polygon"]:
            raise arcpy.ExecuteError("Input must be Polyline or Polygon.")

        if not input_sr or input_sr.name == "Unknown":
            arcpy.AddWarning("Input coordinate system is Unknown. Projection results may not be reliable.")

        if add_projected and projected_sr is None:
            raise arcpy.ExecuteError("Projected coordinate system is required when Add Projected Coordinates is checked.")

        if add_geographic and geographic_sr is None:
            raise arcpy.ExecuteError("Geographic coordinate system is required when Add Geographic Coordinates is checked.")

        # Runtime datum-difference warnings (mirrors validation, in case the
        # tool is called from Python where updateMessages never runs).
        input_gcs = _gcs_name(input_sr)
        if add_projected and input_gcs and _gcs_name(projected_sr) != input_gcs and not projected_transformation:
            arcpy.AddWarning(
                "Datum difference detected for the projected output ({} -> {}) with no "
                "transformation selected. Coordinates may be offset.".format(input_gcs, _gcs_name(projected_sr))
            )
        if add_geographic and input_gcs and _gcs_name(geographic_sr) != input_gcs and not geographic_transformation:
            arcpy.AddWarning(
                "Datum difference detected for the geographic output ({} -> {}) with no "
                "transformation selected. Coordinates may be offset.".format(input_gcs, _gcs_name(geographic_sr))
            )

        # Determine output point geometry spatial reference and its source.
        if output_geometry_sr_option == "Use Selected Projected Coordinate System" and add_projected:
            output_source = "projected"
            output_sr = projected_sr
        elif output_geometry_sr_option == "Use Selected Geographic Coordinate System" and add_geographic:
            output_source = "geographic"
            output_sr = geographic_sr
        else:
            output_source = "input"
            output_sr = input_sr

        out_workspace = os.path.dirname(out_points)
        out_name = os.path.basename(out_points)

        if arcpy.Exists(out_points):
            arcpy.management.Delete(out_points)

        arcpy.management.CreateFeatureclass(
            out_path=out_workspace,
            out_name=out_name,
            geometry_type="POINT",
            spatial_reference=output_sr
        )

        # Core fields.
        arcpy.management.AddField(out_points, "SRC_OID", "LONG")
        arcpy.management.AddField(out_points, "PART_ID", "LONG")
        arcpy.management.AddField(out_points, "VERTEX_ID", "LONG")
        arcpy.management.AddField(out_points, "VERTEX_UID", "TEXT", field_length=80)

        if include_z:
            arcpy.management.AddField(out_points, "Z_VALUE", "DOUBLE")

        # Projected output fields.
        if add_projected:
            arcpy.management.AddField(out_points, "X_PROJ", "DOUBLE")
            arcpy.management.AddField(out_points, "Y_PROJ", "DOUBLE")
            arcpy.management.AddField(out_points, "PROJ_CS", "TEXT", field_length=160)
            arcpy.management.AddField(out_points, "PROJ_WKID", "LONG")

        # Geographic output fields.
        if add_geographic:
            if "Decimal Degrees" in geographic_formats:
                arcpy.management.AddField(out_points, "LONGITUDE", "DOUBLE")
                arcpy.management.AddField(out_points, "LATITUDE", "DOUBLE")

            if "Degrees Minutes Seconds (DMS)" in geographic_formats:
                arcpy.management.AddField(out_points, "LON_DMS", "TEXT", field_length=50)
                arcpy.management.AddField(out_points, "LAT_DMS", "TEXT", field_length=50)

            if "Degrees Decimal Minutes (DDM)" in geographic_formats:
                arcpy.management.AddField(out_points, "LON_DDM", "TEXT", field_length=50)
                arcpy.management.AddField(out_points, "LAT_DDM", "TEXT", field_length=50)

            arcpy.management.AddField(out_points, "GEO_CS", "TEXT", field_length=160)
            arcpy.management.AddField(out_points, "GEO_WKID", "LONG")

        insert_fields = ["SHAPE@XY", "SRC_OID", "PART_ID", "VERTEX_ID", "VERTEX_UID"]

        if include_z:
            insert_fields.append("Z_VALUE")

        if add_projected:
            insert_fields.extend(["X_PROJ", "Y_PROJ", "PROJ_CS", "PROJ_WKID"])

        if add_geographic:
            if "Decimal Degrees" in geographic_formats:
                insert_fields.extend(["LONGITUDE", "LATITUDE"])
            if "Degrees Minutes Seconds (DMS)" in geographic_formats:
                insert_fields.extend(["LON_DMS", "LAT_DMS"])
            if "Degrees Decimal Minutes (DDM)" in geographic_formats:
                insert_fields.extend(["LON_DDM", "LAT_DDM"])
            insert_fields.extend(["GEO_CS", "GEO_WKID"])

        projected_wkid = self._safe_factory_code(projected_sr) if add_projected else 0
        geographic_wkid = self._safe_factory_code(geographic_sr) if add_geographic else 0

        vertex_count = 0
        part_mismatch_warned = False

        with arcpy.da.SearchCursor(in_features, [desc.OIDFieldName, "SHAPE@"]) as search_cursor:
            with arcpy.da.InsertCursor(out_points, insert_fields) as insert_cursor:
                for source_oid, geometry in search_cursor:
                    if geometry is None:
                        continue

                    # Project ONCE per feature (not per vertex) - large speedup
                    # and the place where the datum transformation is applied.
                    projected_geometry = None
                    geographic_geometry = None

                    if add_projected:
                        if projected_transformation:
                            projected_geometry = geometry.projectAs(projected_sr, projected_transformation)
                        else:
                            projected_geometry = geometry.projectAs(projected_sr)

                    if add_geographic:
                        if geographic_transformation:
                            geographic_geometry = geometry.projectAs(geographic_sr, geographic_transformation)
                        else:
                            geographic_geometry = geometry.projectAs(geographic_sr)

                    if output_source == "projected":
                        output_geometry = projected_geometry
                    elif output_source == "geographic":
                        output_geometry = geographic_geometry
                    else:
                        output_geometry = geometry

                    for part_index in range(geometry.partCount):
                        original_points = self._part_points(geometry, part_index)
                        if not original_points:
                            continue

                        projected_points = self._part_points(projected_geometry, part_index) if projected_geometry else None
                        geographic_points = self._part_points(geographic_geometry, part_index) if geographic_geometry else None
                        output_points_list = self._part_points(output_geometry, part_index)

                        # Structural safety check: projection preserves vertex
                        # structure, but guard against any mismatch.
                        lengths = [len(original_points), len(output_points_list)]
                        if projected_points is not None:
                            lengths.append(len(projected_points))
                        if geographic_points is not None:
                            lengths.append(len(geographic_points))
                        n_points = min(lengths)
                        if len(set(lengths)) > 1 and not part_mismatch_warned:
                            arcpy.AddWarning(
                                "Vertex count differs between projected copies of feature {} part {}; "
                                "using the common vertex range.".format(source_oid, part_index + 1)
                            )
                            part_mismatch_warned = True

                        # Optionally drop the duplicated polygon closing vertex
                        # (decided on the ORIGINAL geometry, applied to all copies).
                        if shape_type == "Polygon" and not keep_closing_vertex and n_points > 1:
                            first_point = original_points[0]
                            last_point = original_points[n_points - 1]
                            if abs(first_point.X - last_point.X) < 0.000000001 and abs(first_point.Y - last_point.Y) < 0.000000001:
                                n_points -= 1

                        for vertex_index in range(n_points):
                            original_point = original_points[vertex_index]
                            output_point = output_points_list[vertex_index]

                            row_values = [
                                (output_point.X, output_point.Y),
                                source_oid,
                                part_index + 1,
                                vertex_index + 1,
                                "{}_{}_{}".format(source_oid, part_index + 1, vertex_index + 1)
                            ]

                            if include_z:
                                try:
                                    row_values.append(original_point.Z)
                                except Exception:
                                    row_values.append(None)

                            if add_projected:
                                projected_point = projected_points[vertex_index]
                                row_values.extend([
                                    projected_point.X,
                                    projected_point.Y,
                                    projected_sr.name,
                                    projected_wkid
                                ])

                            if add_geographic:
                                geographic_point = geographic_points[vertex_index]
                                longitude = geographic_point.X
                                latitude = geographic_point.Y

                                if "Decimal Degrees" in geographic_formats:
                                    row_values.extend([longitude, latitude])

                                if "Degrees Minutes Seconds (DMS)" in geographic_formats:
                                    row_values.extend([
                                        self._decimal_degrees_to_dms(longitude, "LON"),
                                        self._decimal_degrees_to_dms(latitude, "LAT")
                                    ])

                                if "Degrees Decimal Minutes (DDM)" in geographic_formats:
                                    row_values.extend([
                                        self._decimal_degrees_to_ddm(longitude, "LON"),
                                        self._decimal_degrees_to_ddm(latitude, "LAT")
                                    ])

                                row_values.extend([
                                    geographic_sr.name,
                                    geographic_wkid
                                ])

                            insert_cursor.insertRow(row_values)
                            vertex_count += 1

        arcpy.AddMessage("Completed successfully.")
        arcpy.AddMessage("Input feature type: {}".format(shape_type))
        arcpy.AddMessage("Total vertices exported: {}".format(vertex_count))
        arcpy.AddMessage("Output point geometry coordinate system option: {}".format(output_geometry_sr_option))

        if add_projected:
            arcpy.AddMessage("Projected coordinate system selected: {}".format(projected_sr.name))
            arcpy.AddMessage("Projected datum transformation: {}".format(projected_transformation or "None"))

        if add_geographic:
            arcpy.AddMessage("Geographic coordinate system selected: {}".format(geographic_sr.name))
            arcpy.AddMessage("Geographic datum transformation: {}".format(geographic_transformation or "None"))
            arcpy.AddMessage("Geographic coordinate format(s): {}".format(", ".join(geographic_formats)))

        return


# -----------------------------------------------------------------------------
# Shared helpers for tools 4-8
# -----------------------------------------------------------------------------

def _parse_angle(text, is_latitude):
    """Parse an angle written as Decimal Degrees, DDM, or DMS into a float.

    Accepts formats such as:
        31.123456            -29.987654
        31 07 24.441 E       N 29 59 12.645
        31d 07' 24.441" E    29:59:12.645 S
        31deg 07.40735' E    (DDM)
    Raises ValueError when the text cannot be interpreted or is out of range.
    """
    if text is None:
        raise ValueError("empty value")
    if isinstance(text, (int, float)):
        value = float(text)
    else:
        s = str(text).strip().upper()
        if not s:
            raise ValueError("empty value")

        hemisphere = None
        for letter in ("N", "S", "E", "W"):
            if s.endswith(letter):
                hemisphere = letter
                s = s[:-1].strip()
                break
            if s.startswith(letter):
                hemisphere = letter
                s = s[1:].strip()
                break

        for token in ("DEG", "MIN", "SEC"):
            s = s.replace(token, " ")
        for ch in ("\u00B0", "\u00BA", "'", "\u2019", '"', "\u201D", "\u2033", "\u2032",
                   ":", ";", ",", "D", "M", "S"):
            s = s.replace(ch, " ")

        parts = [p for p in s.split() if p]
        if not parts:
            raise ValueError("no numeric content in '{}'".format(text))

        try:
            numbers = [float(p) for p in parts]
        except ValueError:
            raise ValueError("non-numeric content in '{}'".format(text))

        negative = numbers[0] < 0 or str(parts[0]).startswith("-")
        degrees = abs(numbers[0])
        minutes = numbers[1] if len(numbers) > 1 else 0.0
        seconds = numbers[2] if len(numbers) > 2 else 0.0

        if minutes < 0 or seconds < 0 or minutes >= 60 or seconds >= 60:
            raise ValueError("minutes/seconds out of range in '{}'".format(text))

        value = degrees + minutes / 60.0 + seconds / 3600.0
        if negative:
            value = -value
        if hemisphere in ("S", "W"):
            value = -abs(value)
        elif hemisphere in ("N", "E"):
            value = abs(value)

    limit = 90.0 if is_latitude else 180.0
    if abs(value) > limit:
        raise ValueError("value {} out of range (+/-{})".format(value, limit))
    return value


def _read_excel_header(excel_path):
    """Return (worksheet_title, header_list) of the preferred sheet."""
    workbook = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    try:
        sheet = None
        if ATTRIBUTES_SHEET in workbook.sheetnames:
            sheet = workbook[ATTRIBUTES_SHEET]
        else:
            for name in workbook.sheetnames:
                candidate = workbook[name]
                if candidate.sheet_state == "visible":
                    sheet = candidate
                    break
            if sheet is None:
                sheet = workbook[workbook.sheetnames[0]]
        header = []
        for row in sheet.iter_rows(min_row=1, max_row=1, values_only=True):
            header = [str(v).strip() if v is not None else "" for v in row]
            break
        return sheet.title, header
    finally:
        workbook.close()


def _infer_field_definition(values):
    """Infer an arcpy field type ('LONG'|'DOUBLE'|'DATE'|'TEXT', length) from sample values."""
    samples = [v for v in values if v is not None and v != ""]
    if not samples:
        return "TEXT", 255

    if all(isinstance(v, (datetime.datetime, datetime.date)) and not isinstance(v, bool) for v in samples):
        return "DATE", 0

    def _is_int(v):
        if isinstance(v, bool):
            return False
        try:
            f = float(v)
            return abs(f - round(f)) < 1e-9 and abs(f) < 2147483647
        except (TypeError, ValueError):
            return False

    def _is_float(v):
        if isinstance(v, bool):
            return False
        try:
            float(v)
            return True
        except (TypeError, ValueError):
            return False

    if all(_is_int(v) for v in samples):
        return "LONG", 0
    if all(_is_float(v) for v in samples):
        return "DOUBLE", 0

    max_length = max(len(str(v)) for v in samples)
    return "TEXT", min(max(max_length + 10, 50), 2000)


def _style_excel_header(worksheet, columns):
    for col_index, name in enumerate(columns, start=1):
        cell = worksheet.cell(row=1, column=col_index, value=name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center")
        worksheet.column_dimensions[get_column_letter(col_index)].width = min(50, max(14, len(name) + 6))
    worksheet.freeze_panes = "A2"


# -----------------------------------------------------------------------------
# Tool 4: Import Points From Excel
# -----------------------------------------------------------------------------
class ImportPointsFromExcel(object):
    COORD_FORMATS = [
        "Auto Detect (DD / DMS / DDM)",
        "Decimal Degrees",
        "Degrees Minutes Seconds (DMS)",
        "Degrees Decimal Minutes (DDM)",
        "Projected XY (numeric)"
    ]

    def __init__(self):
        self.label = "04 - Import Points From Excel"
        self.description = (
            "Create a point feature class from an Excel coordinate list. "
            "Parses Decimal Degrees, DMS (e.g. 31d 07' 24.441\" E), DDM, or "
            "projected XY values, in any coordinate system, with optional "
            "datum transformation. Remaining Excel columns become attribute "
            "fields automatically."
        )
        self.canRunInBackground = False

    def getParameterInfo(self):
        in_excel = arcpy.Parameter(
            displayName="Input Excel File (.xlsx)",
            name="in_excel",
            datatype="DEFile",
            parameterType="Required",
            direction="Input"
        )
        in_excel.filter.list = ["xlsx"]

        x_column = arcpy.Parameter(
            displayName="X / Longitude / Easting Column",
            name="x_column",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        x_column.filter.type = "ValueList"
        x_column.filter.list = []

        y_column = arcpy.Parameter(
            displayName="Y / Latitude / Northing Column",
            name="y_column",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        y_column.filter.type = "ValueList"
        y_column.filter.list = []

        coord_format = arcpy.Parameter(
            displayName="Coordinate Format",
            name="coord_format",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        coord_format.filter.type = "ValueList"
        coord_format.filter.list = self.COORD_FORMATS
        coord_format.value = self.COORD_FORMATS[0]

        input_sr = arcpy.Parameter(
            displayName="Coordinate System of the Excel Values",
            name="input_sr",
            datatype="GPSpatialReference",
            parameterType="Required",
            direction="Input"
        )

        out_points = arcpy.Parameter(
            displayName="Output Point Feature Class",
            name="out_points",
            datatype="DEFeatureClass",
            parameterType="Required",
            direction="Output"
        )

        output_sr = arcpy.Parameter(
            displayName="Output Coordinate System (optional - defaults to input)",
            name="output_sr",
            datatype="GPSpatialReference",
            parameterType="Optional",
            direction="Input"
        )

        transformation = arcpy.Parameter(
            displayName="Datum Transformation (recommended when datums differ)",
            name="transformation",
            datatype="GPString",
            parameterType="Optional",
            direction="Input"
        )
        transformation.filter.type = "ValueList"
        transformation.filter.list = []

        keep_attributes = arcpy.Parameter(
            displayName="Create Attribute Fields From Remaining Columns",
            name="keep_attributes",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        keep_attributes.value = True

        return [in_excel, x_column, y_column, coord_format, input_sr,
                out_points, output_sr, transformation, keep_attributes]

    def updateParameters(self, parameters):
        if parameters[0].value and openpyxl is not None and parameters[0].altered:
            try:
                _, header = _read_excel_header(parameters[0].valueAsText)
                columns = [h for h in header if h]
                parameters[1].filter.list = columns
                parameters[2].filter.list = columns
            except Exception:
                pass

        try:
            if parameters[4].value and parameters[6].value:
                transformations = arcpy.ListTransformations(parameters[4].value, parameters[6].value) or []
                parameters[7].filter.list = transformations
        except Exception:
            pass
        return

    def updateMessages(self, parameters):
        try:
            if parameters[4].value and parameters[6].value and not parameters[7].valueAsText:
                if _gcs_name(parameters[4].value) != _gcs_name(parameters[6].value):
                    parameters[7].setWarningMessage(
                        "The input datum ({}) differs from the output datum ({}). "
                        "Select a datum transformation, or coordinates may be offset.".format(
                            _gcs_name(parameters[4].value), _gcs_name(parameters[6].value))
                    )
        except Exception:
            pass

        fmt = parameters[3].valueAsText or ""
        if parameters[4].value is not None and "Projected" in fmt:
            try:
                if parameters[4].value.type != "Projected":
                    parameters[4].setWarningMessage(
                        "Format is 'Projected XY' but the selected coordinate system is not projected."
                    )
            except Exception:
                pass
        return

    def _parse_pair(self, x_raw, y_raw, coord_format):
        """Return (x, y) as floats according to the selected format."""
        if coord_format == "Projected XY (numeric)":
            return float(x_raw), float(y_raw)
        # DD / DMS / DDM / auto all route through the same tolerant parser.
        x_value = _parse_angle(x_raw, is_latitude=False)
        y_value = _parse_angle(y_raw, is_latitude=True)
        return x_value, y_value

    def execute(self, parameters, messages):
        _require_openpyxl()
        arcpy.env.overwriteOutput = True

        in_excel = parameters[0].valueAsText
        x_column = parameters[1].valueAsText
        y_column = parameters[2].valueAsText
        coord_format = parameters[3].valueAsText
        input_sr = parameters[4].value
        out_points = parameters[5].valueAsText
        output_sr = parameters[6].value
        transformation = (parameters[7].valueAsText or "").strip() or None
        keep_attributes = bool(parameters[8].value)

        if not os.path.exists(in_excel):
            raise arcpy.ExecuteError("Excel file does not exist: {}".format(in_excel))

        workbook = openpyxl.load_workbook(in_excel, data_only=True)
        try:
            if ATTRIBUTES_SHEET in workbook.sheetnames:
                worksheet = workbook[ATTRIBUTES_SHEET]
            else:
                worksheet = None
                for name in workbook.sheetnames:
                    if workbook[name].sheet_state == "visible":
                        worksheet = workbook[name]
                        break
                if worksheet is None:
                    worksheet = workbook[workbook.sheetnames[0]]

            arcpy.AddMessage("Reading worksheet: {}".format(worksheet.title))

            header = [str(c.value).strip() if c.value is not None else "" for c in worksheet[1]]
            header_index = {name.upper(): i for i, name in enumerate(header) if name}

            if x_column.upper() not in header_index:
                raise arcpy.ExecuteError("X column '{}' not found in Excel header.".format(x_column))
            if y_column.upper() not in header_index:
                raise arcpy.ExecuteError("Y column '{}' not found in Excel header.".format(y_column))

            x_index = header_index[x_column.upper()]
            y_index = header_index[y_column.upper()]

            attribute_columns = []
            if keep_attributes:
                attribute_columns = [
                    (name, i) for i, name in enumerate(header)
                    if name and i not in (x_index, y_index)
                ]

            rows = []
            for row in worksheet.iter_rows(min_row=2, values_only=True):
                if row is None:
                    continue
                if all(v is None or v == "" for v in row):
                    continue
                rows.append(row)
        finally:
            workbook.close()

        if not rows:
            raise arcpy.ExecuteError("No data rows found in the Excel file.")

        target_sr = output_sr if output_sr is not None else input_sr
        needs_projection = output_sr is not None

        out_workspace = os.path.dirname(out_points)
        out_name = os.path.basename(out_points)
        if arcpy.Exists(out_points):
            arcpy.management.Delete(out_points)
        arcpy.management.CreateFeatureclass(
            out_path=out_workspace, out_name=out_name,
            geometry_type="POINT", spatial_reference=target_sr
        )

        arcpy.management.AddField(out_points, "SRC_ROW", "LONG")

        field_defs = []
        if attribute_columns:
            used_names = {"SRC_ROW", "OBJECTID", "SHAPE"}
            for column_name, column_index in attribute_columns:
                samples = [r[column_index] if column_index < len(r) else None for r in rows[:200]]
                field_type, field_length = _infer_field_definition(samples)
                base = arcpy.ValidateFieldName(column_name, out_workspace) or "FIELD"
                candidate = base
                suffix = 1
                while candidate.upper() in used_names:
                    suffix += 1
                    candidate = "{}_{}".format(base[:60], suffix)
                used_names.add(candidate.upper())
                if field_type == "TEXT":
                    arcpy.management.AddField(out_points, candidate, "TEXT", field_length=field_length)
                else:
                    arcpy.management.AddField(out_points, candidate, field_type)
                field_defs.append((candidate, column_index, field_type, field_length))

        insert_fields = ["SHAPE@", "SRC_ROW"] + [fd[0] for fd in field_defs]

        imported = 0
        errors = []

        with arcpy.da.InsertCursor(out_points, insert_fields) as cursor:
            for row_number, row in enumerate(rows, start=2):
                x_raw = row[x_index] if x_index < len(row) else None
                y_raw = row[y_index] if y_index < len(row) else None
                try:
                    x_value, y_value = self._parse_pair(x_raw, y_raw, coord_format)
                except (ValueError, TypeError) as parse_error:
                    errors.append((row_number, x_raw, y_raw, str(parse_error)))
                    continue

                point_geometry = arcpy.PointGeometry(arcpy.Point(x_value, y_value), input_sr)
                if needs_projection:
                    if transformation:
                        point_geometry = point_geometry.projectAs(target_sr, transformation)
                    else:
                        point_geometry = point_geometry.projectAs(target_sr)

                values = [point_geometry, row_number]
                for field_name, column_index, field_type, field_length in field_defs:
                    value = row[column_index] if column_index < len(row) else None
                    if value is not None and field_type == "TEXT":
                        value = str(value)[:field_length] if field_length else str(value)
                    values.append(value)

                cursor.insertRow(values)
                imported += 1

        arcpy.AddMessage("Import completed.")
        arcpy.AddMessage("Points imported: {}".format(imported))
        arcpy.AddMessage("Rows skipped: {}".format(len(errors)))
        if errors:
            arcpy.AddWarning("{} row(s) could not be parsed:".format(len(errors)))
            for row_number, x_raw, y_raw, reason in errors[:20]:
                arcpy.AddWarning("  Excel row {} | X='{}' Y='{}' | {}".format(row_number, x_raw, y_raw, reason))
            if len(errors) > 20:
                arcpy.AddWarning("  ... and {} more.".format(len(errors) - 20))
        return


# -----------------------------------------------------------------------------
# Tool 5: Attribute QA/QC Report
# -----------------------------------------------------------------------------
class AttributeQAQCReport(object):
    MAX_ISSUE_ROWS = 50000

    def __init__(self):
        self.label = "05 - Attribute QA/QC Report"
        self.description = (
            "Scan a layer or table for data quality issues - nulls, duplicate "
            "values, leading/trailing/double whitespace, domain violations, and "
            "geometry problems - and write a formatted Excel report."
        )
        self.canRunInBackground = False

    def getParameterInfo(self):
        in_table = arcpy.Parameter(
            displayName="Input Layer or Table",
            name="in_table",
            datatype="GPTableView",
            parameterType="Required",
            direction="Input"
        )

        fields_to_check = arcpy.Parameter(
            displayName="Fields to Check (blank = all attribute fields)",
            name="fields_to_check",
            datatype="Field",
            parameterType="Optional",
            direction="Input",
            multiValue=True
        )
        fields_to_check.parameterDependencies = [in_table.name]

        check_nulls = arcpy.Parameter(
            displayName="Check for Null / Empty Values",
            name="check_nulls",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        check_nulls.value = True

        check_duplicates = arcpy.Parameter(
            displayName="Check for Duplicate Values",
            name="check_duplicates",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        check_duplicates.value = True

        check_whitespace = arcpy.Parameter(
            displayName="Check for Whitespace Problems (text fields)",
            name="check_whitespace",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        check_whitespace.value = True

        check_domains = arcpy.Parameter(
            displayName="Check Domain Violations",
            name="check_domains",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        check_domains.value = True

        check_geometry = arcpy.Parameter(
            displayName="Check Geometry (null / empty / zero size)",
            name="check_geometry",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        check_geometry.value = True

        out_excel = arcpy.Parameter(
            displayName="Output Excel Report (.xlsx)",
            name="out_excel",
            datatype="DEFile",
            parameterType="Required",
            direction="Output"
        )
        out_excel.filter.list = ["xlsx"]

        return [in_table, fields_to_check, check_nulls, check_duplicates,
                check_whitespace, check_domains, check_geometry, out_excel]

    def _load_domains(self, in_table):
        domains = {}
        workspace = _workspace_for_editor(in_table)
        if workspace:
            try:
                for domain in arcpy.da.ListDomains(workspace):
                    domains[domain.name] = domain
            except Exception:
                pass
        return domains

    def execute(self, parameters, messages):
        _require_openpyxl()

        in_table = parameters[0].valueAsText
        selected_fields = _parse_multivalue(parameters[1].valueAsText, [])
        check_nulls = bool(parameters[2].value)
        check_duplicates = bool(parameters[3].value)
        check_whitespace = bool(parameters[4].value)
        check_domains = bool(parameters[5].value)
        check_geometry = bool(parameters[6].value)
        out_excel = parameters[7].valueAsText

        if not out_excel.lower().endswith(".xlsx"):
            out_excel += ".xlsx"

        desc = arcpy.Describe(in_table)
        oid_field = getattr(desc, "OIDFieldName", None)
        has_geometry = hasattr(desc, "shapeType")
        shape_type = getattr(desc, "shapeType", None)

        selected_upper = set(f.upper() for f in selected_fields) if selected_fields else None
        fields = []
        for field in arcpy.ListFields(in_table):
            if field.type in ("OID", "Geometry", "Blob", "Raster", "GlobalID"):
                continue
            if selected_upper and field.name.upper() not in selected_upper:
                continue
            fields.append(field)

        if not fields and not (check_geometry and has_geometry):
            raise arcpy.ExecuteError("No fields to check.")

        domains = self._load_domains(in_table) if check_domains else {}

        cursor_fields = [oid_field] + [f.name for f in fields]
        if check_geometry and has_geometry:
            cursor_fields.append("SHAPE@")

        issues = []          # (oid, field, issue_type, value, detail)
        value_registry = {}  # field -> value -> [oids]
        counters = {}        # (field, issue_type) -> count
        total_rows = 0
        truncated = False

        def _add_issue(oid, field_name, issue_type, value, detail):
            counters[(field_name, issue_type)] = counters.get((field_name, issue_type), 0) + 1
            if len(issues) < self.MAX_ISSUE_ROWS:
                issues.append((oid, field_name, issue_type, value, detail))
            else:
                nonlocal_flag[0] = True

        nonlocal_flag = [False]

        with arcpy.da.SearchCursor(in_table, cursor_fields) as cursor:
            for row in cursor:
                total_rows += 1
                oid = row[0]

                for field_position, field in enumerate(fields, start=1):
                    value = row[field_position]

                    if check_nulls and (value is None or (isinstance(value, str) and value.strip() == "")):
                        _add_issue(oid, field.name, "Null / empty", value, "Value is null or empty")
                        continue

                    if check_whitespace and isinstance(value, str):
                        problems = []
                        if value != value.strip():
                            problems.append("leading/trailing whitespace")
                        if "  " in value:
                            problems.append("double spaces")
                        if problems:
                            _add_issue(oid, field.name, "Whitespace", value, ", ".join(problems))

                    if check_duplicates and value is not None and value != "":
                        registry = value_registry.setdefault(field.name, {})
                        registry.setdefault(value, []).append(oid)

                    if check_domains and field.domain and field.domain in domains and value is not None:
                        domain = domains[field.domain]
                        try:
                            if domain.domainType == "CodedValue":
                                if value not in domain.codedValues:
                                    _add_issue(oid, field.name, "Domain violation", value,
                                               "Not in coded value domain '{}'".format(domain.name))
                            elif domain.domainType == "Range":
                                low, high = domain.range
                                if not (low <= value <= high):
                                    _add_issue(oid, field.name, "Domain violation", value,
                                               "Outside range {} - {} of domain '{}'".format(low, high, domain.name))
                        except Exception:
                            pass

                if check_geometry and has_geometry:
                    geometry = row[-1]
                    if geometry is None:
                        _add_issue(oid, "<geometry>", "Geometry", None, "Null geometry")
                    else:
                        try:
                            if geometry.pointCount == 0:
                                _add_issue(oid, "<geometry>", "Geometry", None, "Empty geometry (0 vertices)")
                            elif shape_type == "Polygon" and geometry.area <= 0:
                                _add_issue(oid, "<geometry>", "Geometry", round(geometry.area, 6), "Zero or negative area")
                            elif shape_type == "Polyline" and geometry.length <= 0:
                                _add_issue(oid, "<geometry>", "Geometry", round(geometry.length, 6), "Zero length")
                        except Exception:
                            pass

        truncated = nonlocal_flag[0]

        if check_duplicates:
            for field_name, registry in value_registry.items():
                for value, oids in registry.items():
                    if len(oids) > 1:
                        counters[(field_name, "Duplicate")] = counters.get((field_name, "Duplicate"), 0) + len(oids)
                        for oid in oids:
                            if len(issues) < self.MAX_ISSUE_ROWS:
                                issues.append((oid, field_name, "Duplicate", value,
                                               "Value shared by {} rows".format(len(oids))))
                            else:
                                truncated = True

        workbook = openpyxl.Workbook()

        summary = workbook.active
        summary.title = "Summary"
        _style_excel_header(summary, ["Field", "Issue Type", "Count"])
        summary_row = 2
        for (field_name, issue_type), count in sorted(counters.items()):
            summary.cell(row=summary_row, column=1, value=field_name)
            summary.cell(row=summary_row, column=2, value=issue_type)
            summary.cell(row=summary_row, column=3, value=count)
            summary_row += 1
        summary.cell(row=summary_row + 1, column=1, value="Rows scanned")
        summary.cell(row=summary_row + 1, column=2, value=total_rows)
        summary.cell(row=summary_row + 2, column=1, value="Source")
        summary.cell(row=summary_row + 2, column=2, value=getattr(desc, "catalogPath", in_table))
        summary.cell(row=summary_row + 3, column=1, value="Generated")
        summary.cell(row=summary_row + 3, column=2, value=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

        detail = workbook.create_sheet("Issues")
        _style_excel_header(detail, [oid_field or "OID", "Field", "Issue Type", "Value", "Detail"])
        for row_index, (oid, field_name, issue_type, value, detail_text) in enumerate(issues, start=2):
            detail.cell(row=row_index, column=1, value=oid)
            detail.cell(row=row_index, column=2, value=field_name)
            detail.cell(row=row_index, column=3, value=issue_type)
            detail.cell(row=row_index, column=4, value=_safe_excel_value(value))
            detail.cell(row=row_index, column=5, value=detail_text)
        detail.auto_filter.ref = detail.dimensions

        folder = os.path.dirname(out_excel)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        workbook.save(out_excel)
        workbook.close()

        arcpy.AddMessage("QA/QC scan completed.")
        arcpy.AddMessage("Rows scanned: {}".format(total_rows))
        arcpy.AddMessage("Issues found: {}".format(sum(counters.values())))
        if truncated:
            arcpy.AddWarning("Issue detail sheet truncated at {} rows; the Summary sheet contains full counts.".format(self.MAX_ISSUE_ROWS))
        arcpy.AddMessage("Report: {}".format(out_excel))
        return


# -----------------------------------------------------------------------------
# Tool 6: Coordinate System Audit
# -----------------------------------------------------------------------------
class CoordinateSystemAudit(object):
    def __init__(self):
        self.label = "06 - Coordinate System Audit"
        self.description = (
            "Walk a workspace (geodatabase or folder) and report the coordinate "
            "system, WKID, datum, and units of every feature class and raster. "
            "Flags Unknown coordinate systems and datum mismatches."
        )
        self.canRunInBackground = False

    def getParameterInfo(self):
        in_workspace = arcpy.Parameter(
            displayName="Workspace to Audit (geodatabase or folder)",
            name="in_workspace",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input"
        )

        out_excel = arcpy.Parameter(
            displayName="Output Excel Report (.xlsx)",
            name="out_excel",
            datatype="DEFile",
            parameterType="Required",
            direction="Output"
        )
        out_excel.filter.list = ["xlsx"]

        return [in_workspace, out_excel]

    def execute(self, parameters, messages):
        _require_openpyxl()

        in_workspace = parameters[0].valueAsText
        out_excel = parameters[1].valueAsText
        if not out_excel.lower().endswith(".xlsx"):
            out_excel += ".xlsx"

        records = []
        datum_counts = {}
        unknown_count = 0

        for dirpath, dirnames, filenames in arcpy.da.Walk(
                in_workspace, datatype=["FeatureClass", "RasterDataset"]):
            for filename in filenames:
                dataset_path = os.path.join(dirpath, filename)
                try:
                    desc = arcpy.Describe(dataset_path)
                    sr = getattr(desc, "spatialReference", None)
                    data_type = getattr(desc, "dataType", "")
                    shape_type = getattr(desc, "shapeType", "") if data_type == "FeatureClass" else ""

                    if sr is None or sr.name == "Unknown":
                        sr_name, wkid, datum, units, sr_type = "Unknown", 0, "Unknown", "", ""
                        unknown_count += 1
                    else:
                        sr_name = sr.name
                        wkid = sr.factoryCode or 0
                        datum = _gcs_name(sr) or ""
                        sr_type = getattr(sr, "type", "")
                        if sr_type == "Projected":
                            units = getattr(sr, "linearUnitName", "")
                        else:
                            units = getattr(sr, "angularUnitName", "")
                        if datum:
                            datum_counts[datum] = datum_counts.get(datum, 0) + 1

                    records.append((
                        os.path.relpath(dataset_path, in_workspace),
                        data_type, shape_type, sr_name, wkid, sr_type, datum, units
                    ))
                except Exception as describe_error:
                    records.append((
                        os.path.relpath(dataset_path, in_workspace),
                        "ERROR", "", str(describe_error)[:100], 0, "", "", ""
                    ))

        if not records:
            raise arcpy.ExecuteError("No feature classes or rasters found in the workspace.")

        dominant_datum = max(datum_counts, key=datum_counts.get) if datum_counts else None

        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "CRS Audit"
        columns = ["Dataset", "Type", "Shape", "Coordinate System", "WKID",
                   "CRS Type", "Datum (GCS)", "Units", "Flag"]
        _style_excel_header(sheet, columns)

        for row_index, record in enumerate(records, start=2):
            flag = ""
            if record[3] == "Unknown":
                flag = "UNKNOWN CRS"
            elif dominant_datum and record[6] and record[6] != dominant_datum:
                flag = "Datum differs from majority ({})".format(dominant_datum)
            for col_index, value in enumerate(record + (flag,), start=1):
                sheet.cell(row=row_index, column=col_index, value=_safe_excel_value(value))
        sheet.auto_filter.ref = sheet.dimensions

        summary = workbook.create_sheet("Summary")
        _style_excel_header(summary, ["Datum", "Dataset Count"])
        for row_index, (datum, count) in enumerate(sorted(datum_counts.items(), key=lambda kv: -kv[1]), start=2):
            summary.cell(row=row_index, column=1, value=datum)
            summary.cell(row=row_index, column=2, value=count)

        folder = os.path.dirname(out_excel)
        if folder and not os.path.exists(folder):
            os.makedirs(folder)
        workbook.save(out_excel)
        workbook.close()

        arcpy.AddMessage("Audit completed: {} dataset(s).".format(len(records)))
        for datum, count in sorted(datum_counts.items(), key=lambda kv: -kv[1]):
            arcpy.AddMessage("  Datum {}: {} dataset(s)".format(datum, count))
        if unknown_count:
            arcpy.AddWarning("{} dataset(s) have an UNKNOWN coordinate system.".format(unknown_count))
        if len(datum_counts) > 1:
            arcpy.AddWarning("Multiple datums detected in this workspace - review the Flag column.")
        arcpy.AddMessage("Report: {}".format(out_excel))
        return


# -----------------------------------------------------------------------------
# Tool 7: Backup Manager
# -----------------------------------------------------------------------------
class BackupManager(object):
    MODES = ["List Backups", "Restore Backup To New Dataset", "Delete Backup"]

    def __init__(self):
        self.label = "07 - Backup Manager"
        self.description = (
            "List, restore, or delete the timestamped backups created by "
            "'02 - Apply Excel Edits Back To Layer'. Restore always copies "
            "to a NEW dataset - it never overwrites existing data."
        )
        self.canRunInBackground = False

    def getParameterInfo(self):
        workspace = arcpy.Parameter(
            displayName="Workspace Containing Backups",
            name="workspace",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input"
        )

        mode = arcpy.Parameter(
            displayName="Action",
            name="mode",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        mode.filter.type = "ValueList"
        mode.filter.list = self.MODES
        mode.value = self.MODES[0]

        backup_name = arcpy.Parameter(
            displayName="Backup Dataset",
            name="backup_name",
            datatype="GPString",
            parameterType="Optional",
            direction="Input"
        )
        backup_name.filter.type = "ValueList"
        backup_name.filter.list = []

        restore_name = arcpy.Parameter(
            displayName="Name for Restored Dataset (Restore mode)",
            name="restore_name",
            datatype="GPString",
            parameterType="Optional",
            direction="Input"
        )

        confirm_delete = arcpy.Parameter(
            displayName="I confirm permanent deletion of the selected backup",
            name="confirm_delete",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        confirm_delete.value = False

        return [workspace, mode, backup_name, restore_name, confirm_delete]

    def _list_backups(self, workspace):
        previous_workspace = arcpy.env.workspace
        try:
            arcpy.env.workspace = workspace
            names = list(arcpy.ListFeatureClasses("*_backup_*") or [])
            names += list(arcpy.ListTables("*_backup_*") or [])
            return sorted(names)
        finally:
            arcpy.env.workspace = previous_workspace

    def updateParameters(self, parameters):
        mode = parameters[1].valueAsText or self.MODES[0]
        parameters[2].enabled = mode != "List Backups"
        parameters[3].enabled = mode == "Restore Backup To New Dataset"
        parameters[4].enabled = mode == "Delete Backup"

        if parameters[0].value and parameters[0].altered:
            try:
                parameters[2].filter.list = self._list_backups(parameters[0].valueAsText)
            except Exception:
                pass
        return

    def updateMessages(self, parameters):
        mode = parameters[1].valueAsText
        if mode in ("Restore Backup To New Dataset", "Delete Backup") and not parameters[2].valueAsText:
            parameters[2].setErrorMessage("Select a backup dataset for this action.")
        if mode == "Restore Backup To New Dataset" and not parameters[3].valueAsText:
            parameters[3].setErrorMessage("Provide a name for the restored dataset.")
        if mode == "Delete Backup" and not bool(parameters[4].value):
            parameters[4].setErrorMessage("Deletion must be explicitly confirmed.")
        return

    @staticmethod
    def _parse_stamp(name):
        try:
            stamp = name.rsplit("_backup_", 1)[1]
            return datetime.datetime.strptime(stamp, "%Y%m%d_%H%M%S").strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return "unknown time"

    def execute(self, parameters, messages):
        workspace = parameters[0].valueAsText
        mode = parameters[1].valueAsText
        backup_name = parameters[2].valueAsText
        restore_name = parameters[3].valueAsText
        confirm_delete = bool(parameters[4].value)

        backups = self._list_backups(workspace)

        if mode == "List Backups":
            if not backups:
                arcpy.AddMessage("No backups found in: {}".format(workspace))
                return
            arcpy.AddMessage("{} backup(s) found in {}:".format(len(backups), workspace))
            for name in backups:
                source = name.rsplit("_backup_", 1)[0]
                arcpy.AddMessage("  {} | source: {} | created: {}".format(name, source, self._parse_stamp(name)))
            return

        if not backup_name:
            raise arcpy.ExecuteError("Select a backup dataset.")
        backup_path = os.path.join(workspace, backup_name)
        if not arcpy.Exists(backup_path):
            raise arcpy.ExecuteError("Backup not found: {}".format(backup_path))

        if mode == "Restore Backup To New Dataset":
            if not restore_name:
                raise arcpy.ExecuteError("Provide a name for the restored dataset.")
            valid_name = arcpy.ValidateTableName(restore_name, workspace)
            restore_path = os.path.join(workspace, valid_name)
            if arcpy.Exists(restore_path):
                raise arcpy.ExecuteError(
                    "A dataset named '{}' already exists. Restore never overwrites - choose another name.".format(valid_name)
                )
            if hasattr(arcpy.Describe(backup_path), "shapeType"):
                arcpy.management.CopyFeatures(backup_path, restore_path)
            else:
                arcpy.management.CopyRows(backup_path, restore_path)
            arcpy.AddMessage("Backup restored to: {}".format(restore_path))
            arcpy.AddMessage(
                "Note: the restore is a NEW dataset. Compare it with your production data "
                "before replacing anything."
            )
            return

        if mode == "Delete Backup":
            if not confirm_delete:
                raise arcpy.ExecuteError("Deletion was not confirmed. Nothing was deleted.")
            arcpy.management.Delete(backup_path)
            arcpy.AddMessage("Backup deleted: {}".format(backup_path))
            return

        raise arcpy.ExecuteError("Unknown action: {}".format(mode))


# -----------------------------------------------------------------------------
# Tool 8: Quick Export For Stakeholders
# -----------------------------------------------------------------------------
class QuickExportForStakeholders(object):
    FORMATS = ["KMZ (Google Earth)", "GeoJSON (WGS84)", "CSV (attributes + WGS84 coordinates)"]

    def __init__(self):
        self.label = "08 - Quick Export For Stakeholders"
        self.description = (
            "One-click export of a layer to KMZ, GeoJSON (WGS84), and/or CSV "
            "for colleagues who do not use GIS."
        )
        self.canRunInBackground = False

    def getParameterInfo(self):
        in_features = arcpy.Parameter(
            displayName="Input Layer",
            name="in_features",
            datatype="GPFeatureLayer",
            parameterType="Required",
            direction="Input"
        )

        formats = arcpy.Parameter(
            displayName="Export Format(s)",
            name="formats",
            datatype="GPString",
            parameterType="Required",
            direction="Input",
            multiValue=True
        )
        formats.filter.type = "ValueList"
        formats.filter.list = self.FORMATS
        formats.values = [[self.FORMATS[0]], [self.FORMATS[1]]]

        out_folder = arcpy.Parameter(
            displayName="Output Folder",
            name="out_folder",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input"
        )

        return [in_features, formats, out_folder]

    def execute(self, parameters, messages):
        in_features = parameters[0].value
        in_features_text = parameters[0].valueAsText
        formats = _parse_multivalue(parameters[1].valueAsText, [])
        out_folder = parameters[2].valueAsText

        desc = arcpy.Describe(in_features_text)
        base_name = os.path.splitext(os.path.basename(desc.name))[0]
        base_name = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in base_name) or "export"

        if not os.path.exists(out_folder):
            os.makedirs(out_folder)

        exported = []

        # A layer object is required for KML export; make one if a path was given.
        layer_for_kml = in_features
        temporary_layer = None
        try:
            if "KMZ (Google Earth)" in formats:
                try:
                    if not hasattr(in_features, "name"):
                        temporary_layer = arcpy.management.MakeFeatureLayer(
                            in_features_text, "quick_export_layer"
                        )[0]
                        layer_for_kml = temporary_layer
                    kmz_path = os.path.join(out_folder, base_name + ".kmz")
                    if os.path.exists(kmz_path):
                        os.remove(kmz_path)
                    arcpy.conversion.LayerToKML(layer_for_kml, kmz_path)
                    exported.append(kmz_path)
                except Exception as kml_error:
                    arcpy.AddWarning("KMZ export failed: {}".format(kml_error))

            if "GeoJSON (WGS84)" in formats:
                try:
                    geojson_path = os.path.join(out_folder, base_name + ".geojson")
                    if os.path.exists(geojson_path):
                        os.remove(geojson_path)
                    arcpy.conversion.FeaturesToJSON(
                        in_features_text, geojson_path,
                        "NOT_FORMATTED", "NO_Z_VALUES", "NO_M_VALUES",
                        "GEOJSON", "WGS84"
                    )
                    exported.append(geojson_path)
                except Exception as geojson_error:
                    arcpy.AddWarning("GeoJSON export failed: {}".format(geojson_error))

            if "CSV (attributes + WGS84 coordinates)" in formats:
                try:
                    csv_path = os.path.join(out_folder, base_name + ".csv")
                    wgs84 = arcpy.SpatialReference(4326)
                    field_names = [
                        f.name for f in arcpy.ListFields(in_features_text)
                        if f.type not in ("Geometry", "Blob", "Raster")
                    ]
                    with open(csv_path, "w", newline="", encoding="utf-8-sig") as csv_file:
                        writer = csv.writer(csv_file)
                        writer.writerow(field_names + ["LONGITUDE_WGS84", "LATITUDE_WGS84"])
                        with arcpy.da.SearchCursor(in_features_text, field_names + ["SHAPE@"]) as cursor:
                            for row in cursor:
                                geometry = row[-1]
                                longitude = latitude = None
                                if geometry is not None:
                                    try:
                                        centroid = geometry.projectAs(wgs84).centroid
                                        longitude = round(centroid.X, 8)
                                        latitude = round(centroid.Y, 8)
                                    except Exception:
                                        pass
                                writer.writerow(list(row[:-1]) + [longitude, latitude])
                    exported.append(csv_path)
                    if desc.shapeType != "Point":
                        arcpy.AddMessage("CSV note: coordinates are feature CENTROIDS for non-point layers.")
                except Exception as csv_error:
                    arcpy.AddWarning("CSV export failed: {}".format(csv_error))
        finally:
            if temporary_layer is not None:
                try:
                    arcpy.management.Delete(temporary_layer)
                except Exception:
                    pass

        if not exported:
            raise arcpy.ExecuteError("No exports were produced. Check the warnings above.")

        arcpy.AddMessage("Export completed. {} file(s):".format(len(exported)))
        for path in exported:
            arcpy.AddMessage("  {}".format(path))
        return

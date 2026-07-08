# -*- coding: utf-8 -*-
"""
Geomatics Combined Tools for ArcGIS Pro
Version: Universal Coordinate System edition

Tools included:
  1. Export Attributes To Excel
  2. Apply Excel Edits Back To Layer
  3. Outline To Vertices Coordinates - Universal Coordinate System Picker

Main vertices-tool change:
  - Removed all preferred / predefined coordinate systems.
  - User selects any Projected Coordinate System using ArcGIS Pro's spatial reference picker.
  - User selects any Geographic Coordinate System using ArcGIS Pro's spatial reference picker.
  - Supports Decimal Degrees, DMS, and DDM geographic output formats.

Prepared for: Mohamed Abdellatief
"""

import arcpy
import os
import math
import datetime

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except Exception:
    openpyxl = None


class Toolbox(object):
    def __init__(self):
        self.label = "Geomatics Combined Tools - Universal CS"
        self.alias = "geomatics_combined_universal_cs"
        self.tools = [
            ExportAttributesToExcel,
            ApplyExcelEditsBackToLayer,
            OutlineToVerticesCoordinatesUniversalCS
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
    desc = arcpy.Describe(dataset)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = os.path.basename(desc.name).replace(".", "_")
    backup_name = "{}_backup_{}".format(base_name, stamp)
    workspace = getattr(desc, "path", None) or arcpy.env.scratchGDB
    backup_path = os.path.join(workspace, backup_name)

    if hasattr(desc, "shapeType"):
        arcpy.management.CopyFeatures(dataset, backup_path)
    else:
        arcpy.management.CopyRows(dataset, backup_path)

    return backup_path


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
        worksheet.title = "Attributes"

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

        metadata = workbook.create_sheet("_GIS_METADATA")
        metadata.sheet_state = "hidden"
        metadata["A1"] = "source"
        metadata["B1"] = getattr(desc, "catalogPath", in_table)
        metadata["A2"] = "key_field"
        metadata["B2"] = key_field
        metadata["A3"] = "instruction"
        metadata["B3"] = "Do not change key values. Edit attributes only. Save and close Excel before applying back."

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
        self.description = "Apply edited XLSX values back to a layer/table using a key field."
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

        workbook = openpyxl.load_workbook(in_excel, data_only=True)
        worksheet = workbook[workbook.sheetnames[0]]

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
                key_value = record.get(excel_key_header)
                if key_value in [None, ""]:
                    continue
                if key_value in excel_edits:
                    duplicate_keys.add(key_value)
                excel_edits[key_value] = record

        workbook.close()

        if duplicate_keys:
            raise arcpy.ExecuteError("Duplicate key value found in Excel: {}".format(list(duplicate_keys)[0]))

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
        matched_rows = 0
        changed_cells = 0
        skipped_blank_cells = 0

        with arcpy.da.UpdateCursor(target_table, cursor_fields) as cursor:
            for db_row in cursor:
                db_key = db_row[0]
                if db_key not in excel_edits:
                    continue

                matched_rows += 1
                excel_row = excel_edits[db_key]
                new_row = list(db_row)
                row_changed = False

                for idx, field_name in enumerate(update_fields, start=1):
                    excel_header_name = header_upper.get(field_name.upper())
                    new_value = excel_row.get(excel_header_name)

                    if new_value in [None, ""]:
                        if not treat_blanks_as_null:
                            skipped_blank_cells += 1
                            continue
                        new_value = None

                    if db_row[idx] != new_value:
                        new_row[idx] = new_value
                        row_changed = True
                        changed_cells += 1

                if row_changed and not dry_run:
                    cursor.updateRow(new_row)

        arcpy.AddMessage("Excel apply process completed.")
        arcpy.AddMessage("Matched rows: {}".format(matched_rows))
        arcpy.AddMessage("Changed cells detected: {}".format(changed_cells))
        arcpy.AddMessage("Blank cells skipped: {}".format(skipped_blank_cells))
        arcpy.AddMessage("Updated fields: {}".format(", ".join(update_fields)))

        if dry_run:
            arcpy.AddWarning("Dry run only: no edits were applied.")
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
            "and/or geographic coordinate system from ArcGIS Pro's coordinate system picker."
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

    def _spatial_reference_kind(self, spatial_reference):
        """Return Geographic, Projected, Unknown, or Other using ArcPy properties where available."""
        if spatial_reference is None:
            return "Unknown"
        try:
            sr_type = spatial_reference.type
            if sr_type:
                return sr_type
        except Exception:
            pass
        try:
            if spatial_reference.factoryCode == 4326 or "GCS" in spatial_reference.exportToString().upper():
                return "Geographic"
        except Exception:
            pass
        return "Other"

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
            add_geographic,              # 4
            geographic_sr,               # 5
            geographic_formats,          # 6
            keep_closing_vertex,         # 7
            include_z,                   # 8
            output_geometry_sr_option    # 9
        ]

    def updateParameters(self, parameters):
        parameters[3].enabled = bool(parameters[2].value)
        parameters[5].enabled = bool(parameters[4].value)
        parameters[6].enabled = bool(parameters[4].value)
        return

    def updateMessages(self, parameters):
        add_projected = bool(parameters[2].value)
        add_geographic = bool(parameters[4].value)

        if not add_projected and not add_geographic:
            parameters[2].setErrorMessage("Select at least one coordinate output: Projected or Geographic.")

        if add_projected and not parameters[3].value:
            parameters[3].setErrorMessage("Please select a projected coordinate system from the coordinate system picker.")

        if add_geographic and not parameters[5].value:
            parameters[5].setErrorMessage("Please select a geographic coordinate system from the coordinate system picker.")

        if add_geographic:
            selected_formats = _parse_multivalue(parameters[6].valueAsText, [])
            if not selected_formats:
                parameters[6].setErrorMessage("Select at least one geographic coordinate number format.")

        option = parameters[9].valueAsText
        if option == "Use Selected Projected Coordinate System" and not add_projected:
            parameters[9].setWarningMessage("Projected output geometry option requires Add Projected Coordinates to be enabled.")
        if option == "Use Selected Geographic Coordinate System" and not add_geographic:
            parameters[9].setWarningMessage("Geographic output geometry option requires Add Geographic Coordinates to be enabled.")

        return

    def execute(self, parameters, messages):
        arcpy.env.overwriteOutput = True

        in_features = parameters[0].valueAsText
        out_points = parameters[1].valueAsText
        add_projected = bool(parameters[2].value)
        projected_sr = parameters[3].value
        add_geographic = bool(parameters[4].value)
        geographic_sr = parameters[5].value
        geographic_formats = _parse_multivalue(parameters[6].valueAsText, ["Decimal Degrees"])
        keep_closing_vertex = bool(parameters[7].value)
        include_z = bool(parameters[8].value)
        output_geometry_sr_option = parameters[9].valueAsText

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

        # Determine output point geometry spatial reference.
        if output_geometry_sr_option == "Use Selected Projected Coordinate System" and add_projected:
            output_sr = projected_sr
        elif output_geometry_sr_option == "Use Selected Geographic Coordinate System" and add_geographic:
            output_sr = geographic_sr
        else:
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

        insert_fields = ["SHAPE@", "SRC_OID", "PART_ID", "VERTEX_ID", "VERTEX_UID"]

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

        with arcpy.da.SearchCursor(in_features, [desc.OIDFieldName, "SHAPE@"] ) as search_cursor:
            with arcpy.da.InsertCursor(out_points, insert_fields) as insert_cursor:
                for source_oid, geometry in search_cursor:
                    if geometry is None:
                        continue

                    for part_index, part in enumerate(geometry):
                        points = [point for point in part if point is not None]

                        if not points:
                            continue

                        if shape_type == "Polygon" and not keep_closing_vertex and len(points) > 1:
                            first_point = points[0]
                            last_point = points[-1]
                            if abs(first_point.X - last_point.X) < 0.000000001 and abs(first_point.Y - last_point.Y) < 0.000000001:
                                points = points[:-1]

                        for vertex_index, point in enumerate(points, start=1):
                            source_point_geometry = arcpy.PointGeometry(point, input_sr)
                            output_point_geometry = source_point_geometry.projectAs(output_sr)

                            row_values = [
                                output_point_geometry,
                                source_oid,
                                part_index + 1,
                                vertex_index,
                                "{}_{}_{}".format(source_oid, part_index + 1, vertex_index)
                            ]

                            if include_z:
                                try:
                                    row_values.append(point.Z)
                                except Exception:
                                    row_values.append(None)

                            if add_projected:
                                projected_point = source_point_geometry.projectAs(projected_sr)
                                row_values.extend([
                                    projected_point.centroid.X,
                                    projected_point.centroid.Y,
                                    projected_sr.name,
                                    projected_wkid
                                ])

                            if add_geographic:
                                geographic_point = source_point_geometry.projectAs(geographic_sr)
                                longitude = geographic_point.centroid.X
                                latitude = geographic_point.centroid.Y

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

        if add_geographic:
            arcpy.AddMessage("Geographic coordinate system selected: {}".format(geographic_sr.name))
            arcpy.AddMessage("Geographic coordinate format(s): {}".format(", ".join(geographic_formats)))

        return

using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.Spreadsheet;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    private static int ConvertExcelPivots(ExcelDocument source, ExcelWorkbookSnapshot snapshot,
        OdsDocument target, ExcelOpenDocumentConversionOptions options,
        IReadOnlyDictionary<string, HashSet<(int Row, int Column)>> convertedCellsBySheet) {
        int converted = 0;
        foreach (ExcelSheet sheet in source.Sheets) {
            foreach (ExcelPivotTableInfo pivot in sheet.GetPivotTables()) {
                if (string.IsNullOrWhiteSpace(pivot.Name) || pivot.SourceSheet == null
                    || pivot.SourceRange == null || pivot.Location == null
                    || !string.Equals(pivot.SourceSheet, pivot.SheetName, StringComparison.Ordinal)
                    || pivot.PageFields.Count != 0 || pivot.DataFields.Count != 1
                    || pivot.RowFields.Count + pivot.ColumnFields.Count == 0
                    || pivot.RowFields.Count + pivot.ColumnFields.Count > 16
                    || pivot.RowGrandTotals == false || pivot.ColumnGrandTotals == false
                    || pivot.Filters.Count != 0 || pivot.CalculatedFields.Count != 0
                    || pivot.Groupings.Count != 0 || pivot.DataFields[0].ShowDataAs != null
                    || pivot.Fields.Any(item => item.HiddenItems.Count > 0
                        || item.SelectedItem != null
                        || item.SubtotalCaption != null || item.InsertPageBreak == true
                        || item.InsertBlankRow == true)
                    || !TryMapPivotFunction(pivot.DataFields[0].Function, out string? function)
                    || !TryGetExcelPivotRanges(pivot, options, snapshot, convertedCellsBySheet,
                        out string? sourceAddress, out string? targetAddress)) continue;
                OdsDataPilotTable authored = target.AddDataPilotTable(pivot.Name, sourceAddress!, targetAddress!);
                foreach (string fieldName in pivot.ColumnFields) authored.AddField(fieldName, "column");
                foreach (string fieldName in pivot.RowFields) authored.AddField(fieldName, "row");
                authored.AddField(pivot.DataFields[0].FieldName, "data", function);
                converted++;
            }
        }
        return converted;
    }

    private static bool TryGetExcelPivotRanges(ExcelPivotTableInfo pivot,
        ExcelOpenDocumentConversionOptions options, ExcelWorkbookSnapshot snapshot,
        IReadOnlyDictionary<string, HashSet<(int Row, int Column)>> convertedCellsBySheet,
        out string? sourceAddress, out string? targetAddress) {
        sourceAddress = targetAddress = null;
        if (!SpreadsheetRangeReference.TryParse(pivot.SourceRange, SpreadsheetAddressDialect.ExcelA1,
                out SpreadsheetRangeReference? source)
            || !SpreadsheetRangeReference.TryParse(pivot.Location, SpreadsheetAddressDialect.ExcelA1,
                out SpreadsheetRangeReference? destination)
            || !source!.IsRange || !destination!.IsRange || !source.Start.IsCell || !source.End!.IsCell
            || !destination.Start.IsCell || !destination.End!.IsCell
            || source.End.Row > options.MaximumRows || source.End.Column > options.MaximumColumns
            || destination.End.Row > options.MaximumRows || destination.End.Column > options.MaximumColumns
            || source.End.Row < source.Start.Row || source.End.Column < source.Start.Column
            || destination.End.Row < destination.Start.Row || destination.End.Column < destination.Start.Column) return false;
        ExcelWorksheetSnapshot? sourceSheet = snapshot.Worksheets.FirstOrDefault(item =>
            string.Equals(item.Name, pivot.SourceSheet, StringComparison.Ordinal));
        if (sourceSheet == null || !convertedCellsBySheet.TryGetValue(pivot.SourceSheet!, out HashSet<(int Row, int Column)>? cells)) return false;
        foreach (ExcelCellSnapshot cell in sourceSheet.Cells) {
            bool withinSource = cell.Row >= source.Start.Row && cell.Row <= source.End.Row
                && cell.Column >= source.Start.Column && cell.Column <= source.End.Column;
            bool withinTarget = cell.Row >= destination.Start.Row && cell.Row <= destination.End.Row
                && cell.Column >= destination.Start.Column && cell.Column <= destination.End.Column;
            if ((withinSource || withinTarget) && !cells.Contains((cell.Row, cell.Column))) return false;
        }
        sourceAddress = SpreadsheetAddressConverter.ExcelRangeToOpenAddress(pivot.SourceRange!, pivot.SourceSheet);
        targetAddress = SpreadsheetAddressConverter.ExcelRangeToOpenAddress(pivot.Location!, pivot.SheetName);
        return sourceAddress.Length > 0 && targetAddress.Length > 0;
    }

    private static bool TryMapPivotFunction(ExcelPivotDataFunction function, out string? result) {
        result = function switch {
            ExcelPivotDataFunction.Sum => "sum",
            ExcelPivotDataFunction.Count => "count",
            ExcelPivotDataFunction.CountNumbers => "countnums",
            ExcelPivotDataFunction.Average => "average",
            ExcelPivotDataFunction.Minimum => "min",
            ExcelPivotDataFunction.Maximum => "max",
            ExcelPivotDataFunction.Product => "product",
            ExcelPivotDataFunction.StandardDeviation => "stdev",
            ExcelPivotDataFunction.StandardDeviationP => "stdevp",
            ExcelPivotDataFunction.Variance => "var",
            ExcelPivotDataFunction.VarianceP => "varp",
            _ => null
        };
        return result != null;
    }

    private static int ConvertOdsDataPilots(OdsDocument source,
        IReadOnlyList<(OdsSheet Source, ExcelSheet Target)> sheets,
        ExcelOpenDocumentConversionOptions options, bool truncated) {
        if (truncated) return 0;
        int converted = 0;
        foreach (OdsDataPilotTable pivot in source.DataPilotTables) {
            if (pivot.HasAdvancedSettings || pivot.Fields.Count == 0 || pivot.Fields.Count > 17
                || !TryGetLocalPivotRanges(pivot, options, out string? sourceSheetName,
                    out string? targetSheetName, out string? sourceRange, out string? destination)) continue;
            (OdsSheet Source, ExcelSheet Target) pair = sheets.FirstOrDefault(item =>
                string.Equals(item.Source.Name, sourceSheetName, StringComparison.Ordinal)
                && string.Equals(item.Source.Name, targetSheetName, StringComparison.Ordinal));
            if (pair.Target == null) continue;
            var rows = new List<string>();
            var columns = new List<string>();
            var data = new List<ExcelPivotDataField>();
            bool unsupported = false;
            foreach (OdsDataPilotField item in pivot.Fields) {
                if (string.IsNullOrWhiteSpace(item.SourceFieldName)) { unsupported = true; break; }
                switch (item.Orientation) {
                    case "row" when item.Function == null || item.Function == "auto":
                        rows.Add(item.SourceFieldName);
                        break;
                    case "column" when item.Function == null || item.Function == "auto":
                        columns.Add(item.SourceFieldName);
                        break;
                    case "data" when TryMapPivotFunction(item.Function, out ExcelPivotDataFunction function):
                        data.Add(new ExcelPivotDataField(item.SourceFieldName, function));
                        break;
                    default:
                        unsupported = true;
                        break;
                }
                if (unsupported) break;
            }
            if (unsupported || data.Count != 1 || rows.Count + columns.Count == 0
                || rows.Count + columns.Count + data.Count != pivot.Fields.Count) continue;
            try {
                pair.Target.AddPivotTable(sourceRange!, destination!, pivot.Name,
                    rowFields: rows, columnFields: columns, dataFields: data,
                    layout: ExcelPivotLayout.Outline);
                converted++;
            } catch (ArgumentException) {
                // Invalid or missing source headers cannot be represented as an Excel pivot.
            } catch (InvalidOperationException) {
                // Existing cell content or pivot layout may prevent safe pivot authoring.
            }
        }
        return converted;
    }

    private static bool TryGetLocalPivotRanges(OdsDataPilotTable pivot,
        ExcelOpenDocumentConversionOptions options, out string? sourceSheetName,
        out string? targetSheetName, out string? sourceRange, out string? destination) {
        sourceSheetName = targetSheetName = sourceRange = destination = null;
        if (!SpreadsheetRangeReference.TryParse(pivot.SourceRangeAddress, SpreadsheetAddressDialect.OpenDocument,
                out SpreadsheetRangeReference? source)
            || !SpreadsheetRangeReference.TryParse(pivot.TargetRangeAddress, SpreadsheetAddressDialect.OpenDocument,
                out SpreadsheetRangeReference? target)
            || !source!.IsRange || !target!.IsRange || !source.Start.IsCell || !source.End!.IsCell
            || !target.Start.IsCell || !target.End!.IsCell) return false;
        sourceSheetName = source.Start.SheetName;
        targetSheetName = target.Start.SheetName;
        if (string.IsNullOrEmpty(sourceSheetName) || string.IsNullOrEmpty(targetSheetName)
            || !string.Equals(sourceSheetName, source.End.SheetName, StringComparison.Ordinal)
            || !string.Equals(targetSheetName, target.End.SheetName, StringComparison.Ordinal)
            || source.End.Row > options.MaximumRows || source.End.Column > options.MaximumColumns
            || target.End.Row > options.MaximumRows || target.End.Column > options.MaximumColumns
            || source.End.Row < source.Start.Row || source.End.Column < source.Start.Column
            || target.End.Row < target.Start.Row || target.End.Column < target.Start.Column) return false;
        sourceRange = SpreadsheetAddressConverter.ToA1((int)source.Start.Row!.Value, source.Start.Column!.Value)
            + ":" + SpreadsheetAddressConverter.ToA1((int)source.End.Row!.Value, source.End.Column!.Value);
        destination = SpreadsheetAddressConverter.ToA1((int)target.Start.Row!.Value, target.Start.Column!.Value);
        return true;
    }

    private static bool TryMapPivotFunction(string? function, out ExcelPivotDataFunction result) {
        switch (function) {
            case "sum": result = ExcelPivotDataFunction.Sum; return true;
            case "count": result = ExcelPivotDataFunction.Count; return true;
            case "countnums": result = ExcelPivotDataFunction.CountNumbers; return true;
            case "average": result = ExcelPivotDataFunction.Average; return true;
            case "min": result = ExcelPivotDataFunction.Minimum; return true;
            case "max": result = ExcelPivotDataFunction.Maximum; return true;
            case "product": result = ExcelPivotDataFunction.Product; return true;
            case "stdev": result = ExcelPivotDataFunction.StandardDeviation; return true;
            case "stdevp": result = ExcelPivotDataFunction.StandardDeviationP; return true;
            case "var": result = ExcelPivotDataFunction.Variance; return true;
            case "varp": result = ExcelPivotDataFunction.VarianceP; return true;
            default: result = default; return false;
        }
    }
}

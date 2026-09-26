using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.Spreadsheet;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    private static int ConvertExcelPivots(ExcelDocument source, ExcelWorkbookSnapshot snapshot,
        OdsDocument target, ExcelOpenDocumentConversionOptions options,
        IReadOnlyDictionary<string, HashSet<(int Row, int Column)>> convertedCellsBySheet) {
        int converted = 0;
        var authoredNames = new HashSet<string>(StringComparer.Ordinal);
        List<ExcelPivotTableInfo> pivots = source.Sheets.SelectMany(sheet => sheet.GetPivotTables()).ToList();
        if (pivots.Count == 0) return 0;
        var neededHeaderRows = new Dictionary<string, HashSet<int>>(StringComparer.Ordinal);
        var neededRanges = new Dictionary<string, List<SpreadsheetRangeReference>>(StringComparer.Ordinal);
        foreach (ExcelPivotTableInfo pivot in pivots) {
            if (pivot.SourceSheet == null
                || !SpreadsheetRangeReference.TryParse(pivot.SourceRange, SpreadsheetAddressDialect.ExcelA1,
                    out SpreadsheetRangeReference? sourceRange) || sourceRange == null || !sourceRange.Start.IsCell) continue;
            if (!neededHeaderRows.TryGetValue(pivot.SourceSheet, out HashSet<int>? rows)) {
                rows = new HashSet<int>();
                neededHeaderRows.Add(pivot.SourceSheet, rows);
            }
            rows.Add((int)sourceRange.Start.Row!.Value);
            foreach (string? address in new[] { pivot.SourceRange, pivot.Location }) {
                if (!SpreadsheetRangeReference.TryParse(address, SpreadsheetAddressDialect.ExcelA1,
                        out SpreadsheetRangeReference? range) || range?.End == null
                    || !range.Start.IsCell || !range.End.IsCell
                    || range.End.Row > options.MaximumRows || range.End.Column > options.MaximumColumns) continue;
                if (!neededRanges.TryGetValue(pivot.SourceSheet, out List<SpreadsheetRangeReference>? ranges)) {
                    ranges = new List<SpreadsheetRangeReference>();
                    neededRanges.Add(pivot.SourceSheet, ranges);
                }
                ranges.Add(range);
            }
        }
        var rangeIndexes = neededRanges.ToDictionary(pair => pair.Key,
            pair => new PivotRangeIndex(pair.Value.Select(range => new PivotRangeIndex.Entry(
                range.Start.Row!.Value, range.End!.Row!.Value, range.Start.Column!.Value, range.End.Column!.Value))),
            StringComparer.Ordinal);
        long remainingIndexWork = 32_000_000;
        var omittedCellsBySheet = new Dictionary<string, List<(int Row, int Column)>>(StringComparer.Ordinal);
        var headersBySheet = new Dictionary<string, Dictionary<int, List<(int Column, string Name)>>>(StringComparer.Ordinal);
        foreach (ExcelWorksheetSnapshot worksheet in snapshot.Worksheets) {
            if (!neededHeaderRows.ContainsKey(worksheet.Name)) continue;
            if (!convertedCellsBySheet.TryGetValue(worksheet.Name, out HashSet<(int Row, int Column)>? convertedCells)) continue;
            ExcelSheet? sourceSheet = source.Sheets.FirstOrDefault(sheet =>
                string.Equals(sheet.Name, worksheet.Name, StringComparison.Ordinal));
            var omitted = new List<(int Row, int Column)>();
            var headerRows = new Dictionary<int, List<(int Column, string Name)>>();
            neededHeaderRows.TryGetValue(worksheet.Name, out HashSet<int>? neededRows);
            rangeIndexes.TryGetValue(worksheet.Name, out PivotRangeIndex? rangeIndex);
            foreach (ExcelCellSnapshot cell in worksheet.Cells) {
                if (rangeIndex == null) continue;
                bool relevant = rangeIndex.Intersects(cell.Row, cell.Row, cell.Column, cell.Column, ref remainingIndexWork);
                if (remainingIndexWork <= 0) return 0;
                if (!relevant) continue;
                if (!convertedCells.Contains((cell.Row, cell.Column))) omitted.Add((cell.Row, cell.Column));
                if (neededRows?.Contains(cell.Row) == true && sourceSheet?.TryGetCellText(cell.Row, cell.Column, out string? text) == true
                    && !string.IsNullOrWhiteSpace(text)) {
                    if (!headerRows.TryGetValue(cell.Row, out List<(int Column, string Name)>? headers)) {
                        headers = new List<(int Column, string Name)>();
                        headerRows.Add(cell.Row, headers);
                    }
                    headers.Add((cell.Column, text.Trim()));
                }
            }
            omitted.Sort((left, right) => left.Row != right.Row
                ? left.Row.CompareTo(right.Row) : left.Column.CompareTo(right.Column));
            omittedCellsBySheet.Add(worksheet.Name, omitted);
            headersBySheet.Add(worksheet.Name, headerRows);
        }
        foreach (ExcelPivotTableInfo pivot in pivots) {
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
                    || pivot.HasValuesAxisField
                    || pivot.RowFields.Concat(pivot.ColumnFields).Distinct(StringComparer.OrdinalIgnoreCase).Count()
                        != pivot.RowFields.Count + pivot.ColumnFields.Count
                    || !TryMapPivotFunction(pivot.DataFields[0].Function, out string? function)
                    || authoredNames.Contains(pivot.Name)
                    || !TryGetExcelPivotRanges(pivot, options, snapshot, target, omittedCellsBySheet, headersBySheet,
                        out string? sourceAddress, out string? targetAddress)) continue;
                OdsDataPilotTable authored = target.AddDataPilotTable(pivot.Name, sourceAddress!, targetAddress!);
                foreach (string fieldName in pivot.ColumnFields) authored.AddField(fieldName, "column");
                foreach (string fieldName in pivot.RowFields) authored.AddField(fieldName, "row");
                authored.AddField(pivot.DataFields[0].FieldName, "data", function);
                authoredNames.Add(pivot.Name);
                converted++;
        }
        return converted;
    }

    private static bool TryGetExcelPivotRanges(ExcelPivotTableInfo pivot,
        ExcelOpenDocumentConversionOptions options, ExcelWorkbookSnapshot snapshot, OdsDocument target,
        IReadOnlyDictionary<string, List<(int Row, int Column)>> omittedCellsBySheet,
        IReadOnlyDictionary<string, Dictionary<int, List<(int Column, string Name)>>> headersBySheet,
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
        if (sourceSheet == null || !omittedCellsBySheet.TryGetValue(pivot.SourceSheet!, out List<(int Row, int Column)>? omitted)) return false;
        if (!headersBySheet.TryGetValue(pivot.SourceSheet!, out Dictionary<int, List<(int Column, string Name)>>? headerRows)
            || !headerRows.TryGetValue((int)source.Start.Row!.Value, out List<(int Column, string Name)>? headerCells)) return false;
        var headerCounts = headerCells
            .Where(cell => cell.Column >= source.Start.Column && cell.Column <= source.End.Column)
            .GroupBy(cell => cell.Name, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        if (pivot.RowFields.Concat(pivot.ColumnFields).Append(pivot.DataFields[0].FieldName)
            .Any(field => !headerCounts.TryGetValue(field, out int count) || count != 1)) return false;
        int low = 0, high = omitted.Count;
        long firstRow = Math.Min(source.Start.Row!.Value, destination.Start.Row!.Value);
        long lastRow = Math.Max(source.End.Row!.Value, destination.End.Row!.Value);
        while (low < high) {
            int middle = low + (high - low) / 2;
            if (omitted[middle].Row < firstRow) low = middle + 1;
            else high = middle;
        }
        for (int index = low; index < omitted.Count && omitted[index].Row <= lastRow; index++) {
            (int row, int column) = omitted[index];
            bool withinSource = row >= source.Start.Row && row <= source.End.Row
                && column >= source.Start.Column && column <= source.End.Column;
            bool withinTarget = row >= destination.Start.Row && row <= destination.End.Row
                && column >= destination.Start.Column && column <= destination.End.Column;
            if (withinSource || withinTarget) return false;
        }
        OdsSheet? projectedSheet = target.GetSheet(pivot.SourceSheet!);
        if (projectedSheet == null || headerCells.Any(cell =>
            cell.Column >= source.Start.Column && cell.Column <= source.End.Column
            && !string.Equals(projectedSheet.Cell(source.Start.Row.Value - 1, cell.Column - 1).Text,
                cell.Name, StringComparison.Ordinal))) return false;
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
        ExcelOpenDocumentConversionOptions options,
        IReadOnlyDictionary<string, List<long>> convertedCellsBySheet) {
        int converted = 0;
        // The Excel cache scans the source for each pivot. Bound the total work across the conversion.
        long remainingPivotScanCells = 1_000_000;
        var convertedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var generatedEntries = new Dictionary<string, List<PivotRangeIndex.Entry>>(StringComparer.Ordinal);
        IReadOnlyList<OdsDataPilotTable> pivots = source.DataPilotTables;
        foreach (OdsDataPilotTable pivot in pivots) {
            if (!TryGetLocalPivotRanges(pivot, options, out _, out string? sheet, out _, out _, out long row, out long column)
                || sheet == null) continue;
            int width = pivot.Fields.Count(field => field.Orientation is "row" or "column" or "data");
            if (width == 0) continue;
            if (!generatedEntries.TryGetValue(sheet, out List<PivotRangeIndex.Entry>? entries)) {
                entries = new List<PivotRangeIndex.Entry>(); generatedEntries.Add(sheet, entries);
            }
            entries.Add(new PivotRangeIndex.Entry(row, row + 1, column, column + width - 1, pivot));
        }
        var generatedIndexes = generatedEntries.ToDictionary(pair => pair.Key,
            pair => new PivotRangeIndex(pair.Value), StringComparer.Ordinal);
        long remainingCollisionWork = 1_000_000;
        var sourceFootprints = new List<(string Sheet, long FirstRow, long LastRow, long FirstColumn, long LastColumn)>();
        // Reserve every retained local source before any output is authored, including later pivots.
        foreach (OdsDataPilotTable pivot in pivots) {
            if (SpreadsheetRangeReference.TryParse(pivot.SourceRangeAddress, SpreadsheetAddressDialect.OpenDocument,
                    out SpreadsheetRangeReference? bounds) && bounds?.End != null
                && bounds.Start.IsCell && bounds.End.IsCell && bounds.Start.SheetName != null
                && string.Equals(bounds.Start.SheetName, bounds.End.SheetName, StringComparison.Ordinal)) {
                sourceFootprints.Add((bounds.Start.SheetName, bounds.Start.Row!.Value, bounds.End.Row!.Value,
                    bounds.Start.Column!.Value, bounds.End.Column!.Value));
            }
        }
        var sourceIndexes = sourceFootprints.GroupBy(item => item.Sheet, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => new PivotRangeIndex(group.Select(item =>
                new PivotRangeIndex.Entry(item.FirstRow, item.LastRow, item.FirstColumn, item.LastColumn))), StringComparer.Ordinal);
        var rowRunsBySheet = new Dictionary<OdsSheet, IReadOnlyList<OdsRowRun>>();
        foreach (OdsDataPilotTable pivot in pivots) {
            if (pivot.HasAdvancedSettings || pivot.Fields.Count == 0 || pivot.Fields.Count > 17
                || string.IsNullOrWhiteSpace(pivot.Name) || pivot.Name != pivot.Name.Trim()
                || convertedNames.Contains(pivot.Name)
                || !TryGetLocalPivotRanges(pivot, options, out string? sourceSheetName,
                    out string? targetSheetName, out string? sourceRange, out string? destination,
                    out long destinationRow, out long destinationColumn)) continue;
            if (!SpreadsheetRangeReference.TryParse(pivot.SourceRangeAddress, SpreadsheetAddressDialect.OpenDocument,
                    out SpreadsheetRangeReference? sourceBounds) || sourceBounds?.End == null
                || !SpreadsheetRangeReference.TryParse(pivot.TargetRangeAddress, SpreadsheetAddressDialect.OpenDocument,
                    out SpreadsheetRangeReference? targetBounds) || targetBounds?.End == null) continue;
            long sourceRows = sourceBounds.End.Row!.Value - sourceBounds.Start.Row!.Value + 1;
            long sourceColumns = sourceBounds.End.Column!.Value - sourceBounds.Start.Column!.Value + 1;
            long targetRows = targetBounds.End.Row!.Value - targetBounds.Start.Row!.Value + 1;
            long targetColumns = targetBounds.End.Column!.Value - targetBounds.Start.Column!.Value + 1;
            if (sourceRows <= 0 || sourceColumns <= 0 || sourceRows > remainingPivotScanCells / sourceColumns)
                continue;
            long sourceArea = sourceRows * sourceColumns;
            if (targetRows <= 0 || targetColumns <= 0 || targetRows > (remainingPivotScanCells - sourceArea) / targetColumns)
                continue;
            long targetArea = targetRows * targetColumns;
            (OdsSheet Source, ExcelSheet Target) pair = sheets.FirstOrDefault(item =>
                string.Equals(item.Source.Name, sourceSheetName, StringComparison.Ordinal)
                && string.Equals(item.Source.Name, targetSheetName, StringComparison.Ordinal));
            if (pair.Target == null ||
                !convertedCellsBySheet.TryGetValue(pair.Source.Name, out List<long>? convertedCells)) continue;
            if (!rowRunsBySheet.TryGetValue(pair.Source, out IReadOnlyList<OdsRowRun>? rowRuns)) {
                rowRuns = pair.Source.RowRuns;
                rowRunsBySheet.Add(pair.Source, rowRuns);
            }
            // Debit only scans actually attempted; a rejected range must not consume the next pivot's budget.
            remainingPivotScanCells -= sourceArea;
            if (!PivotRangeRetained(rowRuns, pivot.SourceRangeAddress!, convertedCells)) continue;
            remainingPivotScanCells -= targetArea;
            if (!PivotRangeRetained(rowRuns, pivot.TargetRangeAddress!, convertedCells)) continue;
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
                || rows.Count + columns.Count + data.Count != pivot.Fields.Count
                || rows.Concat(columns).Distinct(StringComparer.OrdinalIgnoreCase).Count() != rows.Count + columns.Count
                || HasAmbiguousPivotHeaders(rowRuns, sourceBounds, pivot.Fields, pair.Target)) continue;
            long generatedLastRow = destinationRow + 1;
            long generatedLastColumn = destinationColumn + rows.Count + columns.Count + data.Count - 1;
            if (generatedLastRow > Math.Min(options.MaximumRows, 1_048_576)
                || generatedLastColumn > Math.Min(options.MaximumColumns, 16_384)
                || generatedLastRow > targetBounds.End.Row!.Value
                || generatedLastColumn > targetBounds.End.Column!.Value
                || PivotFootprintsOverlap(sourceBounds.Start.Row!.Value, sourceBounds.End.Row!.Value,
                    sourceBounds.Start.Column!.Value, sourceBounds.End.Column!.Value,
                    destinationRow, generatedLastRow, destinationColumn, generatedLastColumn)
                || (generatedIndexes.TryGetValue(pair.Source.Name, out PivotRangeIndex? generatedIndex)
                    && (generatedIndex.Intersects(destinationRow, generatedLastRow, destinationColumn, generatedLastColumn, ref remainingCollisionWork)
                        || generatedIndex.Intersects(sourceBounds.Start.Row!.Value, sourceBounds.End.Row!.Value,
                            sourceBounds.Start.Column!.Value, sourceBounds.End.Column!.Value, ref remainingCollisionWork)))
                || (sourceIndexes.TryGetValue(pair.Source.Name, out PivotRangeIndex? sourceIndex)
                    && sourceIndex.Intersects(destinationRow, generatedLastRow, destinationColumn, generatedLastColumn, ref remainingCollisionWork))) continue;
            try {
                pair.Target.AddPivotTable(sourceRange!, destination!, pivot.Name,
                    rowFields: rows, columnFields: columns, dataFields: data,
                    layout: ExcelPivotLayout.Outline);
                convertedNames.Add(pivot.Name);
                generatedIndexes[pair.Source.Name].Enable(pivot);
                converted++;
            } catch (ArgumentException) {
                // Invalid or missing source headers cannot be represented as an Excel pivot.
            } catch (InvalidOperationException) {
                // Existing cell content or pivot layout may prevent safe pivot authoring.
            }
        }
        return converted;
    }

    private static bool PivotFootprintsOverlap(long firstRow, long lastRow, long firstColumn, long lastColumn,
        long otherFirstRow, long otherLastRow, long otherFirstColumn, long otherLastColumn) =>
        firstRow <= otherLastRow && lastRow >= otherFirstRow &&
        firstColumn <= otherLastColumn && lastColumn >= otherFirstColumn;

    private static bool HasAmbiguousPivotHeaders(IReadOnlyList<OdsRowRun> rowRuns, SpreadsheetRangeReference range,
        IReadOnlyList<OdsDataPilotField> fields, ExcelSheet target) {
        long firstRow = range.Start.Row!.Value - 1;
        long firstColumn = range.Start.Column!.Value - 1;
        long lastColumn = range.End!.Column!.Value - 1;
        OdsRowRun? header = rowRuns.FirstOrDefault(run =>
            firstRow >= run.StartRow && firstRow < SaturatingAdd(run.StartRow, run.RepeatCount));
        if (header == null) return true;
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var exactNames = new HashSet<string>(StringComparer.Ordinal);
        long nextColumn = firstColumn;
        foreach (OdsCellRun cell in header.CellRuns) {
            if (cell.StartColumn > lastColumn) break;
            long end = Math.Min(lastColumn, SaturatingAdd(cell.StartColumn, cell.RepeatCount) - 1);
            if (end < firstColumn) continue;
            long start = Math.Max(cell.StartColumn, firstColumn);
            if (cell.IsCovered || start != nextColumn) return true;
            OdsCellValue value = cell.Value;
            string name = value.DisplayText;
            if (value.Kind == OdsCellValueKind.Empty || string.IsNullOrWhiteSpace(name)
                || name != name.Trim() || !names.Add(name) || end > start
                || !target.TryGetCellText((int)(firstRow + 1), (int)(start + 1), out string projectedName)
                || !string.Equals(projectedName, name, StringComparison.Ordinal))
                return true;
            exactNames.Add(name);
            nextColumn = end + 1;
        }
        return nextColumn <= lastColumn || fields.Any(field => !exactNames.Contains(field.SourceFieldName));
    }

    private static bool PivotRangeRetained(IReadOnlyList<OdsRowRun> rowRuns, string address,
        List<long> convertedCells) {
        if (!SpreadsheetRangeReference.TryParse(address, SpreadsheetAddressDialect.OpenDocument,
                out SpreadsheetRangeReference? range) || !range!.IsRange || !range.Start.IsCell || !range.End!.IsCell) return false;
        long firstRow = range.Start.Row!.Value, lastRow = range.End.Row!.Value;
        long firstColumn = range.Start.Column!.Value, lastColumn = range.End.Column!.Value;
        foreach (OdsRowRun rowRun in rowRuns) {
            long runFirstRow = rowRun.StartRow + 1;
            long runLastRow = SaturatingAdd(rowRun.StartRow, rowRun.RepeatCount);
            if (runFirstRow > lastRow) break;
            if (runLastRow < firstRow) continue;
            foreach (OdsCellRun cellRun in rowRun.CellRuns) {
                if (cellRun.IsCovered || !IsSignificant(cellRun)) continue;
                long runFirstColumn = cellRun.StartColumn + 1;
                long runLastColumn = SaturatingAdd(cellRun.StartColumn, cellRun.RepeatCount);
                if (runFirstColumn > lastColumn) break;
                if (runLastColumn < firstColumn) continue;
                for (long row = Math.Max(firstRow, runFirstRow); row <= Math.Min(lastRow, runLastRow); row++) {
                    for (long column = Math.Max(firstColumn, runFirstColumn); column <= Math.Min(lastColumn, runLastColumn); column++) {
                        if (convertedCells.BinarySearch((row << 15) | (uint)column) < 0) return false;
                    }
                }
            }
        }
        return true;
    }

    private static bool TryGetLocalPivotRanges(OdsDataPilotTable pivot,
        ExcelOpenDocumentConversionOptions options, out string? sourceSheetName,
        out string? targetSheetName, out string? sourceRange, out string? destination,
        out long destinationRow, out long destinationColumn) {
        sourceSheetName = targetSheetName = sourceRange = destination = null;
        destinationRow = destinationColumn = 0;
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
        destinationRow = target.Start.Row!.Value;
        destinationColumn = target.Start.Column!.Value;
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

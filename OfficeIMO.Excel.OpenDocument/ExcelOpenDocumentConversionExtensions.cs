using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.Spreadsheet;
using System.Text;

namespace OfficeIMO.Excel.OpenDocument;

/// <summary>Explicit conversions between OfficeIMO Excel and native OpenDocument spreadsheet models.</summary>
public static partial class ExcelOpenDocumentConversionExtensions {
    /// <summary>Converts an Excel workbook to an in-memory ODS document.</summary>
    public static OdsDocument ToOpenDocument(this ExcelDocument source,
        ExcelOpenDocumentConversionOptions? options = null) => source.ToOpenDocumentResult(options).Value;

    /// <summary>Converts an Excel workbook to an in-memory ODS document and reports every lossy mapping.</summary>
    public static OdfConversionResult<OdsDocument> ToOpenDocumentResult(this ExcelDocument source,
        ExcelOpenDocumentConversionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        ExcelOpenDocumentConversionOptions effective = options ?? new ExcelOpenDocumentConversionOptions();
        effective.Validate();
        ExcelWorkbookSnapshot snapshot = source.CreateInspectionSnapshot();
        OdsDocument target = OdsDocument.Create();
        var report = new OdfConversionReport(source.SourceFormat.ToString().ToUpperInvariant(), "ODS");
        target.Metadata.Title = snapshot.Title;

        int cells = 0, formulas = 0, formulaTranslationFailures = 0, styles = 0, hyperlinks = 0, comments = 0, threadedComments = 0, merges = 0;
        int rows = 0, columns = 0, convertedValidations = 0, skippedValidations = 0, tables = 0, filters = 0, unsupportedStyles = 0, skippedStyles = 0;
        long materializedCells = 0, skippedCells = 0, skippedRows = 0, skippedColumns = 0, skippedMerges = 0;
        bool truncated = false;
        var dataStyles = new Dictionary<uint, string>();
        int worksheetOrdinal = 0;
        foreach (ExcelWorksheetSnapshot worksheet in snapshot.Worksheets) {
            worksheetOrdinal++;
            OdsSheet sheet = target.AddSheet(worksheet.Name);
            var materializedCoordinates = new HashSet<(int Row, int Column)>();
            sheet.Hidden = worksheet.Hidden;
            foreach (ExcelColumnSnapshot column in worksheet.Columns) {
                if (column.EndIndex > effective.MaximumColumns) {
                    skippedColumns += Math.Max(0, column.EndIndex - Math.Max(effective.MaximumColumns + 1, column.StartIndex) + 1L);
                    truncated = true;
                }
                int last = Math.Min(column.EndIndex, effective.MaximumColumns);
                for (int index = Math.Max(1, column.StartIndex); index <= last; index++) {
                    OdsColumn converted = sheet.Column(index - 1L);
                    converted.Hidden = column.Hidden;
                    if (column.Width.HasValue) converted.Width = OdfLength.Points(ExcelWidthToPoints(column.Width.Value));
                    columns++;
                }
            }
            foreach (ExcelRowSnapshot row in worksheet.Rows) {
                if (row.Index < 1 || row.Index > effective.MaximumRows) { skippedRows++; truncated = true; continue; }
                OdsRow converted = sheet.Row(row.Index - 1L);
                converted.Hidden = row.Hidden;
                if (row.Height.HasValue) converted.Height = OdfLength.Points(row.Height.Value);
                rows++;
            }

            foreach (ExcelCellSnapshot cell in worksheet.Cells) {
                if (cell.Row < 1 || cell.Column < 1 || cell.Row > effective.MaximumRows || cell.Column > effective.MaximumColumns ||
                    materializedCells >= effective.MaximumExpandedCells) {
                    skippedCells++;
                    truncated = true;
                    continue;
                }
                materializedCells++;
                materializedCoordinates.Add((cell.Row, cell.Column));
                OdsCell converted = sheet.Cell(cell.Row - 1L, cell.Column - 1L);
                if (!string.IsNullOrWhiteSpace(cell.Formula)) {
                    var translation = SpreadsheetAddressConverter.ExcelFormulaToOpenFormula(cell.Formula!);
                    if (translation.IsSuccessful) {
                        converted.Formula = translation.Formula;
                        formulas++;
                    } else {
                        formulaTranslationFailures++;
                    }
                }
                bool exactValue = SetOdsValue(converted, cell.Value);
                if (!exactValue) unsupportedStyles++;
                if (cell.Hyperlink != null && !string.IsNullOrWhiteSpace(cell.Hyperlink.Target)) {
                    string href = cell.Hyperlink.IsExternal
                        ? cell.Hyperlink.Target
                        : "#" + SpreadsheetAddressConverter.ExcelRangeToOpenAddress(cell.Hyperlink.Target);
                    converted.SetHyperlink(ValueText(cell.Value), href);
                    hyperlinks++;
                }
                if (effective.IncludeBasicStyles && cell.Style != null) {
                    ApplyExcelStyle(target, converted, cell.Style, dataStyles, ref unsupportedStyles);
                    styles++;
                } else if (cell.Style != null) {
                    skippedStyles++;
                }
                if (cell.Comment != null) {
                    converted.AddAnnotation(cell.Comment.Text, cell.Comment.Author);
                    comments++;
                }
                cells++;
            }

            var threadedAnnotations = new Dictionary<(int Row, int Column), List<ExcelThreadedCommentSnapshot>>();
            foreach (ExcelThreadedCommentSnapshot threaded in worksheet.ThreadedComments) {
                if (!ExcelReference.TryParse(threaded.CellReference, out ExcelReference? reference)
                    || reference!.Kind != ExcelReferenceKind.Cell
                    || reference.IsQualified
                    || reference.Start.Row < 1
                    || reference.Start.Column < 1
                    || reference.Start.Row > effective.MaximumRows
                    || reference.Start.Column > effective.MaximumColumns) {
                    skippedCells++;
                    truncated = true;
                    continue;
                }
                var coordinate = (reference.Start.Row, reference.Start.Column);
                if (!threadedAnnotations.TryGetValue(coordinate, out List<ExcelThreadedCommentSnapshot>? annotationThread)) {
                    if (!materializedCoordinates.Contains(coordinate)) {
                        if (materializedCells >= effective.MaximumExpandedCells) {
                            skippedCells++;
                            truncated = true;
                            continue;
                        }
                        materializedCoordinates.Add(coordinate);
                        materializedCells++;
                    }
                    annotationThread = new List<ExcelThreadedCommentSnapshot>();
                    threadedAnnotations.Add(coordinate, annotationThread);
                }
                annotationThread.Add(threaded);
                threadedComments++;
            }
            foreach (KeyValuePair<(int Row, int Column), List<ExcelThreadedCommentSnapshot>> entry in threadedAnnotations
                .OrderBy(item => item.Key.Row).ThenBy(item => item.Key.Column)) {
                OdsCell cell = sheet.Cell(entry.Key.Row - 1L, entry.Key.Column - 1L);
                ExcelThreadedCommentSnapshot first = entry.Value[0];
                string text = FormatThreadedCommentTranscript(entry.Value, includeMetadataForSingleRoot: cell.Annotation != null);
                if (cell.Annotation != null) {
                    cell.Annotation.Text = cell.Annotation.Text + "\n\nThreaded discussion:\n" + text;
                } else {
                    DateTimeOffset? date = first.Date.HasValue
                        ? new DateTimeOffset(first.Date.Value.ToUniversalTime())
                        : (DateTimeOffset?)null;
                    cell.AddAnnotation(text, first.Author, date, first.Id);
                }
            }

            int validationOrdinal = 0;
            foreach (ExcelDataValidationSnapshot validation in worksheet.Validations) {
                validationOrdinal++;
                if (!TryCreateOdsValidationCondition(validation, out OdsValidationConditionSyntax? condition)) {
                    skippedValidations++;
                    continue;
                }

                string validationName = "validation_" + worksheetOrdinal.ToString(CultureInfo.InvariantCulture)
                    + "_" + validationOrdinal.ToString(CultureInfo.InvariantCulture);
                bool assigned = false;
                bool validationIncomplete = false;
                bool validationLimitReached = false;
                foreach (string a1Range in validation.A1Ranges) {
                    if (validationLimitReached) break;
                    if (!SpreadsheetRangeReference.TryParse(a1Range, SpreadsheetAddressDialect.ExcelA1, out SpreadsheetRangeReference? parsed)
                        || !parsed!.Start.IsCell || parsed.Start.SheetName != null
                        || (parsed.End != null && (!parsed.End.IsCell || parsed.End.SheetName != null))) {
                        validationIncomplete = true;
                        continue;
                    }
                    SpreadsheetCellReference end = parsed.End ?? parsed.Start;
                    long firstRow = parsed.Start.Row!.Value;
                    int firstColumn = parsed.Start.Column!.Value;
                    long lastRow = end.Row!.Value;
                    int lastColumn = end.Column!.Value;
                    if (firstRow > lastRow || firstColumn > lastColumn
                        || firstRow > effective.MaximumRows || firstColumn > effective.MaximumColumns) {
                        validationIncomplete = true;
                        truncated = true;
                        continue;
                    }
                    if (lastRow > effective.MaximumRows || lastColumn > effective.MaximumColumns) {
                        validationIncomplete = true;
                        truncated = true;
                    }
                    lastRow = Math.Min(lastRow, effective.MaximumRows);
                    lastColumn = Math.Min(lastColumn, effective.MaximumColumns);
                    for (long row = firstRow; row <= lastRow; row++) {
                        for (int column = firstColumn; column <= lastColumn; column++) {
                            var coordinate = (checked((int)row), column);
                            if (!materializedCoordinates.Contains(coordinate)) {
                                if (materializedCells >= effective.MaximumExpandedCells) {
                                    truncated = true;
                                    validationIncomplete = true;
                                    validationLimitReached = true;
                                    break;
                                }
                                materializedCoordinates.Add(coordinate);
                                materializedCells++;
                            }
                            sheet.Cell(row - 1L, column - 1L).ValidationName = validationName;
                            assigned = true;
                        }
                        if (validationLimitReached) break;
                    }
                }
                if (assigned) {
                    OdsValidation convertedValidation = target.AddValidation(validationName, condition!, validation.AllowBlank);
                    if (validation.PromptTitle != null || validation.Prompt != null) {
                        convertedValidation.SetHelpMessage(validation.PromptTitle, validation.Prompt, validation.ShowInputMessage);
                    }
                    if (validation.ErrorTitle != null || validation.Error != null) {
                        convertedValidation.SetErrorMessage(
                            validation.ErrorTitle,
                            validation.Error,
                            ParseOdsValidationMessageType(validation.ErrorStyle),
                            validation.ShowErrorMessage);
                    }
                    convertedValidations++;
                }
                if (!assigned || validationIncomplete) {
                    skippedValidations++;
                }
            }

            foreach (ExcelMergedRangeSnapshot merged in worksheet.MergedRanges) {
                if (merged.StartRow < 1 || merged.StartColumn < 1 || merged.StartRow > effective.MaximumRows ||
                    merged.StartColumn > effective.MaximumColumns || merged.EndRow > effective.MaximumRows ||
                    merged.EndColumn > effective.MaximumColumns) {
                    skippedMerges++;
                    truncated = true;
                    continue;
                }
                long rowSpan = merged.EndRow - merged.StartRow + 1L;
                long columnSpan = merged.EndColumn - merged.StartColumn + 1L;
                long mergeCells = checked(rowSpan * columnSpan);
                if (mergeCells > OdsSheet.DefaultMaximumMergeCells) {
                    skippedMerges++;
                    truncated = true;
                    continue;
                }
                long remaining = effective.MaximumExpandedCells - materializedCells;
                long newlyMaterializedCells = 0;
                for (int row = merged.StartRow; row <= merged.EndRow && newlyMaterializedCells <= remaining; row++) {
                    for (int column = merged.StartColumn; column <= merged.EndColumn; column++) {
                        if (!materializedCoordinates.Contains((row, column))) newlyMaterializedCells++;
                        if (newlyMaterializedCells > remaining) break;
                    }
                }
                if (newlyMaterializedCells > remaining) {
                    skippedMerges++;
                    truncated = true;
                    continue;
                }
                sheet.Merge(merged.StartRow - 1L, merged.StartColumn - 1L, rowSpan, columnSpan, mergeCells);
                for (int row = merged.StartRow; row <= merged.EndRow; row++) {
                    for (int column = merged.StartColumn; column <= merged.EndColumn; column++) {
                        materializedCoordinates.Add((row, column));
                    }
                }
                materializedCells += newlyMaterializedCells;
                merges++;
            }
            tables += worksheet.Tables.Count;
            if (worksheet.AutoFilter != null) filters++;
            if (worksheet.FrozenRowCount > 0 || worksheet.FrozenColumnCount > 0 || worksheet.RightToLeft || !worksheet.ShowGridlines) {
                report.Add("worksheet-views", OdfConversionMappingStatus.Unsupported, 1,
                    "Frozen panes and Excel-specific worksheet view settings are not represented by the current ODS typed surface.");
            }
            if (worksheet.Protection != null) report.Add("worksheet-protection", OdfConversionMappingStatus.Unsupported, 1);
        }

        int namedRanges = 0, builtInNames = 0, disambiguatedNames = 0;
        var usedNamedRanges = new HashSet<string>(StringComparer.Ordinal);
        foreach (ExcelNamedRangeSnapshot named in snapshot.NamedRanges) {
            if (named.IsBuiltIn) { builtInNames++; continue; }
            string address = SpreadsheetAddressConverter.ExcelRangeToOpenAddress(named.ReferenceA1, named.SheetName);
            if (address.Length == 0) continue;
            string outputName = named.Name;
            if (!usedNamedRanges.Add(outputName)) {
                outputName = CreateUniqueNamedRangeName(named.Name, named.SheetName, usedNamedRanges);
                disambiguatedNames++;
            }
            target.AddNamedRange(outputName, address);
            namedRanges++;
        }

        AddConverted(report, "worksheets", snapshot.Worksheets.Count);
        AddConverted(report, "cells", cells);
        AddConverted(report, "rows", rows);
        if (columns > 0) report.Add("column-layout", OdfConversionMappingStatus.Approximated, columns,
            "Excel character-unit column widths are converted to approximate physical widths.");
        AddConverted(report, "merges", merges);
        AddConverted(report, "hyperlinks", hyperlinks);
        AddConverted(report, "named-ranges", namedRanges);
        if (disambiguatedNames > 0) report.Add("sheet-local-named-ranges", OdfConversionMappingStatus.Approximated, disambiguatedNames,
            "Duplicate sheet-local Excel names were made unique because ODS named ranges are workbook scoped.");
        if (formulas > 0) report.Add("formulas", OdfConversionMappingStatus.Approximated, formulas,
            "Formula syntax is parsed and translated to OpenFormula; cached values are retained.");
        if (formulaTranslationFailures > 0) report.Add("formulas", OdfConversionMappingStatus.Skipped, formulaTranslationFailures,
            "Formula syntax without a safe OpenFormula representation was omitted; the cached cell value was retained.");
        if (styles > 0) report.Add("cell-styles", OdfConversionMappingStatus.Approximated, styles,
            "Bold, italic, font, foreground, fill, and common number formats are mapped; other Excel style details are omitted.");
        if (skippedStyles > 0) report.Add("cell-styles", OdfConversionMappingStatus.Skipped, skippedStyles,
            "Cell styles were omitted because IncludeBasicStyles is disabled.");
        if (unsupportedStyles > 0) report.Add("cell-format-details", OdfConversionMappingStatus.Unsupported, unsupportedStyles);
        AddConverted(report, "comments", comments);
        if (threadedComments > 0) report.Add("threaded-comments", OdfConversionMappingStatus.Approximated, threadedComments,
            "Each cell thread was flattened into one schema-valid ODS annotation transcript that retains comment bodies and available author, timestamp, identity, parent, and resolved-state metadata.");
        AddConverted(report, "validations", convertedValidations);
        if (skippedValidations > 0) report.Add("validations", OdfConversionMappingStatus.Unsupported, skippedValidations,
            "Unsupported validation rules and ranges, plus assignments clipped by configured conversion limits, were not mapped completely.");
        AddUnsupported(report, "structured-tables", tables, "Table cells remain; Excel table semantics and styles are not translated.");
        AddUnsupported(report, "filters", filters, "Filter state is not translated.");
        AddUnsupported(report, "built-in-names", builtInNames, "Excel print-area and print-title names are not translated.");
        AddUnsupported(report, "charts", snapshot.ChartPartCount, "Excel chart parts are not translated to ODS.");
        AddUnsupported(report, "pivot-tables", snapshot.PivotTablePartCount, "Excel pivot-table parts are not translated to ODS.");
        AddUnsupported(report, "slicers", snapshot.SlicerPartCount, null);
        AddUnsupported(report, "timelines", snapshot.TimelinePartCount, null);
        AddUnsupported(report, "slicer-binding-metadata", snapshot.SlicerBindingMetadataPartCount,
            "OfficeIMO slicer binding metadata is not represented in ODS.");
        AddUnsupported(report, "timeline-binding-metadata", snapshot.TimelineBindingMetadataPartCount,
            "OfficeIMO timeline binding metadata is not represented in ODS.");
        AddUnsupported(report, "connections", snapshot.ConnectionPartCount, null);
        AddUnsupported(report, "query-tables", snapshot.QueryTablePartCount, null);
        if (truncated) report.Add("expansion-limits", OdfConversionMappingStatus.Skipped, 1,
            $"Configured limits omitted content or assignments, including {skippedCells} cells, {skippedRows} rows, {skippedColumns} columns, and {skippedMerges} merges.");
        return new OdfConversionResult<OdsDocument>(target, report).ApplyPolicy(effective.LossPolicy);
    }

    private static string CreateUniqueNamedRangeName(string name, string? sheetName, HashSet<string> usedNames) {
        string suffix = new string((sheetName ?? "Sheet").Select(character => char.IsLetterOrDigit(character) ? character : '_').ToArray());
        if (suffix.Length == 0) suffix = "Sheet";
        string candidate = name + "__" + suffix;
        int index = 2;
        while (!usedNames.Add(candidate)) candidate = name + "__" + suffix + "_" + index++.ToString(CultureInfo.InvariantCulture);
        return candidate;
    }

    /// <summary>Converts an ODS document to an in-memory Excel workbook.</summary>
    public static ExcelDocument ToExcelDocument(this OdsDocument source,
        ExcelOpenDocumentConversionOptions? options = null) => source.ToExcelDocumentResult(options).Value;

    /// <summary>Converts an ODS document to an in-memory Excel workbook and reports every lossy mapping.</summary>
    public static OdfConversionResult<ExcelDocument> ToExcelDocumentResult(this OdsDocument source,
        ExcelOpenDocumentConversionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        ExcelOpenDocumentConversionOptions effective = options ?? new ExcelOpenDocumentConversionOptions();
        effective.Validate();
        ExcelDocument target = ExcelDocument.Create(new MemoryStream());
        var report = new OdfConversionReport("ODS", "XLSX");
        target.BuiltinDocumentProperties.Title = source.Metadata.Title;
        var dataStyles = source.DataStyles.GroupBy(style => style.Name, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        var sourceValidations = source.Validations
            .GroupBy(validation => validation.Name, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        var sourceNamedRangeNames = new HashSet<string>(
            source.NamedRanges.Select(static namedRange => namedRange.Name),
            StringComparer.Ordinal);

        long expandedCells = 0;
        int cells = 0, formulas = 0, formulaTranslationFailures = 0, styles = 0, hyperlinks = 0, externalHyperlinks = 0, comments = 0, combinedComments = 0, metadataTranscriptComments = 0, merges = 0, rowLayouts = 0, columnLayouts = 0;
        int invalidValues = 0, validations = 0, convertedValidations = 0, unsupportedValidationAssignments = 0, unsupportedHyperlinks = 0, unsupportedMeasurements = 0, unsupportedDataStyleFormats = 0, skippedStyles = 0, renamedSheets = 0, worksheetCount = 0;
        int forcedVisibleWorksheets = 0;
        bool truncated = false;
        ExcelSheet? activeTarget = null;
        ExcelSheet? firstTarget = null;
        foreach (OdsSheet odsSheet in source.Sheets) {
            ExcelSheet sheet = target.AddWorksheet(odsSheet.Name);
            var validationTargets = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            firstTarget ??= sheet;
            worksheetCount++;
            if (!string.Equals(sheet.Name, odsSheet.Name, StringComparison.Ordinal)) renamedSheets++;
            sheet.SetHidden(odsSheet.Hidden);
            if (!odsSheet.Hidden && activeTarget == null) activeTarget = sheet;

            foreach (OdsColumnRun columnRun in odsSheet.ColumnRuns) {
                long columnEnd = SaturatingAdd(columnRun.StartColumn, columnRun.RepeatCount);
                long lastExclusive = Math.Min(columnEnd, effective.MaximumColumns);
                for (long column = columnRun.StartColumn; column < lastExclusive; column++) {
                    if (!columnRun.Hidden && !columnRun.Width.HasValue) continue;
                    int excelColumn = checked((int)column + 1);
                    if (columnRun.Hidden) sheet.SetColumnHidden(excelColumn, true);
                    if (columnRun.Width.HasValue) {
                        if (columnRun.Width.Value.TryToPoints(out double points)) sheet.SetColumnWidth(excelColumn, PointsToExcelWidth(points));
                        else unsupportedMeasurements++;
                    }
                    columnLayouts++;
                }
                if (columnEnd > effective.MaximumColumns) truncated = true;
            }

            foreach (OdsRowRun rowRun in odsSheet.RowRuns) {
                long rowEnd = SaturatingAdd(rowRun.StartRow, rowRun.RepeatCount);
                long lastRowExclusive = Math.Min(rowEnd, effective.MaximumRows);
                if (rowEnd > effective.MaximumRows) truncated = true;
                for (long row = rowRun.StartRow; row < lastRowExclusive; row++) {
                    int excelRow = checked((int)row + 1);
                    if (rowRun.Hidden) { sheet.SetRowHidden(excelRow, true); rowLayouts++; }
                    if (rowRun.Height.HasValue) {
                        if (rowRun.Height.Value.TryToPoints(out double points)) sheet.SetRowHeight(excelRow, points);
                        else unsupportedMeasurements++;
                        rowLayouts++;
                    }

                    foreach (OdsCellRun cellRun in rowRun.CellRuns) {
                        long cellColumnEnd = SaturatingAdd(cellRun.StartColumn, cellRun.RepeatCount);
                        long lastColumnExclusive = Math.Min(cellColumnEnd, effective.MaximumColumns);
                        if (cellColumnEnd > effective.MaximumColumns) truncated = true;
                        if (cellRun.IsCovered || !IsSignificant(cellRun)) continue;
                        for (long column = cellRun.StartColumn; column < lastColumnExclusive; column++) {
                            if (expandedCells >= effective.MaximumExpandedCells) { truncated = true; break; }
                            expandedCells++;
                            int excelColumn = checked((int)column + 1);
                            ExcelCell converted = sheet.CellAt(excelRow, excelColumn);
                            if (!SetExcelValue(converted, cellRun.Value)) invalidValues++;
                            if (!string.IsNullOrWhiteSpace(cellRun.Formula)) {
                                var translation = SpreadsheetAddressConverter.OpenFormulaToExcel(cellRun.Formula!);
                                if (translation.IsSuccessful) {
                                    converted.SetFormula(translation.Formula);
                                    formulas++;
                                } else {
                                    formulaTranslationFailures++;
                                }
                            }
                            if (!string.IsNullOrWhiteSpace(cellRun.HyperlinkHref)) {
                                string href = cellRun.HyperlinkHref!;
                                bool convertedHyperlink = false;
                                if (href.StartsWith("#", StringComparison.Ordinal)) {
                                    string fragment = href.Substring(1);
                                    string location = SpreadsheetAddressConverter.OpenAddressToExcel(fragment);
                                    if (location.Length == 0 && sourceNamedRangeNames.Contains(fragment)) location = fragment;
                                    if (location.Length > 0) {
                                        sheet.SetInternalLink(excelRow, excelColumn, location, cellRun.Text, style: true);
                                        convertedHyperlink = true;
                                    } else {
                                        unsupportedHyperlinks++;
                                    }
                                } else {
                                    sheet.SetHyperlink(excelRow, excelColumn, href, cellRun.Text, style: true);
                                    convertedHyperlink = true;
                                }
                                if (cellRun.Value.Kind != OdsCellValueKind.Empty) _ = SetExcelValue(converted, cellRun.Value);
                                if (!string.IsNullOrWhiteSpace(cellRun.Formula)) {
                                    var translation = SpreadsheetAddressConverter.OpenFormulaToExcel(cellRun.Formula!);
                                    if (translation.IsSuccessful) converted.SetFormula(translation.Formula);
                                }
                                if (convertedHyperlink) {
                                    hyperlinks++;
                                    if (IsExternalOdfHref(href)) externalHyperlinks++;
                                }
                            }
                            if (effective.IncludeBasicStyles && cellRun.StyleName != null) {
                                unsupportedMeasurements += ApplyOdsStyle(converted, cellRun, dataStyles,
                                    out bool unsupportedDataStyleFormat);
                                if (unsupportedDataStyleFormat) unsupportedDataStyleFormats++;
                                styles++;
                            } else if (cellRun.StyleName != null) {
                                skippedStyles++;
                            }
                            if (cellRun.ValidationName != null) {
                                validations++;
                                if (!validationTargets.TryGetValue(cellRun.ValidationName, out List<string>? targets)) {
                                    targets = new List<string>();
                                    validationTargets.Add(cellRun.ValidationName, targets);
                                }
                                targets.Add(SpreadsheetAddressConverter.ToA1(excelRow, excelColumn));
                            }
                            if (cellRun.Annotations.Count > 0) {
                                OdsAnnotation first = cellRun.Annotations[0];
                                bool preserveSingleMetadata = cellRun.Annotations.Count == 1
                                    && (!string.IsNullOrWhiteSpace(first.Name) || first.Date.HasValue);
                                string commentText = cellRun.Annotations.Count == 1 && !preserveSingleMetadata
                                    ? first.Text
                                    : string.Join("\n\n", cellRun.Annotations.Select(FormatAnnotationForExcel));
                                sheet.SetComment(excelRow, excelColumn, commentText,
                                    string.IsNullOrWhiteSpace(first.Creator) ? "OfficeIMO" : first.Creator!);
                                comments += cellRun.Annotations.Count;
                                if (cellRun.Annotations.Count > 1) combinedComments += cellRun.Annotations.Count;
                                else if (preserveSingleMetadata) metadataTranscriptComments++;
                            }
                            cells++;

                            if (cellRun.RowSpan > 1 || cellRun.ColumnSpan > 1) {
                                long mergeLastRow = SaturatingAdd(row, cellRun.RowSpan);
                                long mergeLastColumn = SaturatingAdd(column, cellRun.ColumnSpan);
                                if (mergeLastRow <= effective.MaximumRows && mergeLastColumn <= effective.MaximumColumns) {
                                    string start = SpreadsheetAddressConverter.ToA1(excelRow, excelColumn);
                                    string end = SpreadsheetAddressConverter.ToA1(
                                        checked((int)mergeLastRow), checked((int)mergeLastColumn));
                                    sheet.MergeRange(start + ":" + end);
                                    merges++;
                                } else truncated = true;
                            }
                        }
                        if (expandedCells >= effective.MaximumExpandedCells) break;
                    }
                    if (expandedCells >= effective.MaximumExpandedCells) break;
                }
                if (expandedCells >= effective.MaximumExpandedCells) break;
            }

            foreach (KeyValuePair<string, List<string>> entry in validationTargets) {
                string references = string.Join(" ", entry.Value.Distinct(StringComparer.Ordinal));
                if (!sourceValidations.TryGetValue(entry.Key, out OdsValidation? validation)
                    || !TryApplyOdsValidation(sheet, references, validation)) {
                    unsupportedValidationAssignments += entry.Value.Count;
                    continue;
                }
                ApplyOdsValidationMessages(sheet, entry.Value[0], validation);
                convertedValidations++;
            }
        }

        if (target.Sheets.Count == 0) activeTarget = target.AddWorksheet("Sheet1");
        else if (activeTarget == null) {
            activeTarget = firstTarget!;
            activeTarget.SetHidden(false);
            forcedVisibleWorksheets++;
        }
        if (activeTarget != null) target.SetActiveWorksheet(activeTarget);

        int namedRanges = 0;
        foreach (OdsNamedRange named in source.NamedRanges) {
            string reference = SpreadsheetAddressConverter.OpenAddressToExcel(named.CellRangeAddress);
            if (reference.Length == 0) continue;
            target.SetNamedRange(named.Name, reference, save: false);
            namedRanges++;
        }

        AddConverted(report, "worksheets", worksheetCount);
        AddConverted(report, "cells", cells);
        AddConverted(report, "row-layout", rowLayouts);
        if (columnLayouts > 0) report.Add("column-layout", OdfConversionMappingStatus.Approximated, columnLayouts,
            "Physical ODF column widths are converted to approximate Excel character widths.");
        AddConverted(report, "merges", merges);
        AddConverted(report, "hyperlinks", hyperlinks);
        AddUnsupported(report, "hyperlinks", unsupportedHyperlinks,
            "Internal hyperlink fragments that are neither cell addresses nor transferred named ranges were omitted.");
        AddConverted(report, "comments", comments - combinedComments - metadataTranscriptComments);
        if (combinedComments > 0) report.Add("comments", OdfConversionMappingStatus.Approximated, combinedComments,
            "Multiple annotations on one ODS cell were combined into one Excel legacy comment.");
        if (metadataTranscriptComments > 0) report.Add("comments", OdfConversionMappingStatus.Approximated,
            metadataTranscriptComments,
            "ODS annotation timestamps and stable names were retained in the Excel legacy comment transcript because legacy comments have no equivalent metadata fields.");
        AddConverted(report, "named-ranges", namedRanges);
        if (formulas > 0) report.Add("formulas", OdfConversionMappingStatus.Approximated, formulas,
            "OpenFormula syntax is parsed and translated to Excel A1 syntax; cached ODS values remain independently represented.");
        if (formulaTranslationFailures > 0) report.Add("formulas", OdfConversionMappingStatus.Skipped, formulaTranslationFailures,
            "Formula syntax without a safe Excel A1 representation was omitted; the cached cell value was retained.");
        if (styles > 0) report.Add("cell-styles", OdfConversionMappingStatus.Approximated, styles,
            "Basic font, fill, and data-style categories are mapped.");
        if (skippedStyles > 0) report.Add("cell-styles", OdfConversionMappingStatus.Skipped, skippedStyles,
            "Cell styles were omitted because IncludeBasicStyles is disabled.");
        if (renamedSheets > 0) report.Add("worksheet-names", OdfConversionMappingStatus.Approximated, renamedSheets,
            "Worksheet names that are not valid in XLSX were sanitized; formulas and named-range text may still use the source names.");
        if (forcedVisibleWorksheets > 0) report.Add("worksheet-visibility", OdfConversionMappingStatus.Approximated,
            forcedVisibleWorksheets, "The first worksheet was made visible because XLSX requires at least one visible worksheet.");
        AddConverted(report, "validations", convertedValidations);
        if (unsupportedValidationAssignments > 0) report.Add("validations", OdfConversionMappingStatus.Unsupported,
            unsupportedValidationAssignments,
            "Only explicit lists and scalar whole-number, decimal, and text-length ODF validation conditions have an exact Excel mapping.");
        AddUnsupported(report, "invalid-values", invalidValues, "Invalid typed lexemes were transferred as display text.");
        AddUnsupported(report, "relative-measurements", unsupportedMeasurements,
            "Relative or unsupported ODF row, column, or text measurements could not be projected to fixed Excel sizes and were omitted.");
        AddUnsupported(report, "cell-format-details", unsupportedDataStyleFormats,
            $"ODF data styles that exceed Excel's {OdsDataStyle.MaximumExcelNumberFormatCodeLength}-character custom number-format limit were omitted.");
        if (truncated) report.Add("expansion-limits", OdfConversionMappingStatus.Skipped, 1,
            "Content outside the configured row, column, or expanded-cell limits was not materialized.");
        AddUnmappedOdfFindings(source.InspectFeatures(), report, formulas, convertedValidations,
            externalHyperlinks, comments, namedRanges);
        target = Normalize(target);
        return new OdfConversionResult<ExcelDocument>(target, report).ApplyPolicy(effective.LossPolicy);
    }

    private static bool IsSignificant(OdsCellRun cell) => cell.Value.Kind != OdsCellValueKind.Empty ||
        cell.Formula != null || cell.StyleName != null || cell.ValidationName != null || cell.HyperlinkHref != null ||
        cell.Annotations.Count > 0 || cell.RowSpan > 1 || cell.ColumnSpan > 1;

    private static string FormatAnnotationForExcel(OdsAnnotation annotation) {
        var header = new List<string>();
        if (!string.IsNullOrWhiteSpace(annotation.Creator)) header.Add(annotation.Creator!);
        if (annotation.Date.HasValue) header.Add(annotation.Date.Value.ToString("u", CultureInfo.InvariantCulture));
        if (!string.IsNullOrWhiteSpace(annotation.Name)) header.Add("Id: " + annotation.Name);
        return header.Count == 0 ? annotation.Text : "[" + string.Join(" — ", header) + "]\n" + annotation.Text;
    }

    private static string FormatThreadedCommentTranscript(
        IReadOnlyList<ExcelThreadedCommentSnapshot> comments,
        bool includeMetadataForSingleRoot = false) {
        if (!includeMetadataForSingleRoot && comments.Count == 1
            && string.IsNullOrWhiteSpace(comments[0].ParentId) && !comments[0].Done) {
            return comments[0].Text;
        }
        var builder = new StringBuilder();
        for (int index = 0; index < comments.Count; index++) {
            ExcelThreadedCommentSnapshot comment = comments[index];
            if (index > 0) builder.Append("\n\n");
            builder.Append(string.IsNullOrWhiteSpace(comment.ParentId) ? "Comment" : "Reply");
            if (!string.IsNullOrWhiteSpace(comment.Author)) builder.Append(" by ").Append(comment.Author);
            if (comment.Date.HasValue) builder.Append(" — ").Append(comment.Date.Value.ToUniversalTime().ToString("u", CultureInfo.InvariantCulture));
            if (comment.Done) builder.Append(" — resolved");
            if (!string.IsNullOrWhiteSpace(comment.Id)) builder.Append("\nId: ").Append(comment.Id);
            if (!string.IsNullOrWhiteSpace(comment.ParentId)) builder.Append("\nParent: ").Append(comment.ParentId);
            builder.Append('\n').Append(comment.Text);
        }
        return builder.ToString();
    }

    private static bool SetOdsValue(OdsCell target, object? value) {
        if (value == null) return true;
        if (value is string text) target.SetString(text);
        else if (value is bool boolean) target.SetBoolean(boolean);
        else if (value is decimal decimalValue) target.SetDecimal(decimalValue);
        else if (value is DateTime dateTime) target.SetDate(dateTime);
        else if (value is DateTimeOffset dateTimeOffset) target.SetDateTime(dateTimeOffset);
        else if (value is TimeSpan timeSpan) target.SetDuration(timeSpan);
        else if (IsNumeric(value)) target.SetNumber(Convert.ToDouble(value, CultureInfo.InvariantCulture));
        else { target.SetString(Convert.ToString(value, CultureInfo.InvariantCulture)); return false; }
        return true;
    }

    private static bool SetExcelValue(ExcelCell target, OdsCellValue value) {
        try {
            switch (value.Kind) {
                case OdsCellValueKind.Empty: return true;
                case OdsCellValueKind.String: target.SetValue(value.LexicalValue); return true;
                case OdsCellValueKind.Number:
                case OdsCellValueKind.Percentage:
                case OdsCellValueKind.Currency: target.SetValue(value.AsDecimal()); return true;
                case OdsCellValueKind.Boolean: target.SetValue(value.AsBoolean()); return true;
                case OdsCellValueKind.Date: target.SetValue(value.AsDateTimeOffset()); return true;
                case OdsCellValueKind.Time: target.SetValue(value.AsTimeSpan()); return true;
                default: target.SetValue(value.ToString()); return false;
            }
        } catch (FormatException) {
            target.SetValue(value.ToString());
            return false;
        } catch (OverflowException) {
            target.SetValue(value.ToString());
            return false;
        }
    }

    private static void ApplyExcelStyle(OdsDocument document, OdsCell target, ExcelCellStyleSnapshot style,
        IDictionary<uint, string> dataStyles, ref int unsupported) {
        if (style.Bold) target.Bold = true;
        if (style.Italic) target.Italic = true;
        if (style.FontSize.HasValue) target.FontSize = OdfLength.Points(style.FontSize.Value);
        if (!string.IsNullOrWhiteSpace(style.FontName)) target.FontFamily = style.FontName;
        if (!string.IsNullOrWhiteSpace(style.FontColorHex)) target.Color = OdfColor.Parse(style.FontColorHex!);
        if (!string.IsNullOrWhiteSpace(style.FillColorHex)) target.BackgroundColor = OdfColor.Parse(style.FillColorHex!);
        if (!string.IsNullOrWhiteSpace(style.NumberFormatCode) && style.NumberFormatCode != "General") {
            if (!dataStyles.TryGetValue(style.StyleIndex, out string? name)) {
                name = "xlData" + style.StyleIndex.ToString(CultureInfo.InvariantCulture);
                SpreadsheetNumberFormatSyntax format = SpreadsheetNumberFormatSyntax.Parse(style.NumberFormatCode!);
                if (style.IsDateLike) {
                    bool timeOnly = format.Tokens.Any(token => token.Kind == SpreadsheetNumberFormatTokenKind.DateTimeSymbol &&
                            (token.Value.IndexOf("h", StringComparison.OrdinalIgnoreCase) >= 0 || token.Value.IndexOf("s", StringComparison.OrdinalIgnoreCase) >= 0)) &&
                        !format.Tokens.Any(token => token.Kind == SpreadsheetNumberFormatTokenKind.DateTimeSymbol &&
                            (token.Value.IndexOf("y", StringComparison.OrdinalIgnoreCase) >= 0 || token.Value.IndexOf("d", StringComparison.OrdinalIgnoreCase) >= 0));
                    if (timeOnly) document.AddTimeStyle(name); else document.AddDateStyle(name);
                } else if (format.IsPercentage) {
                    document.AddPercentageStyle(name, format.DecimalPlaces, format.UsesGrouping);
                } else if (!string.IsNullOrWhiteSpace(format.CurrencySymbol)) {
                    document.AddCurrencyStyle(name, format.CurrencySymbol!, format.DecimalPlaces, format.UsesGrouping);
                } else {
                    document.AddNumberStyle(name, format.DecimalPlaces, format.UsesGrouping);
                }
                if (!format.IsValid || format.SectionCount > 1 || format.Tokens.Any(token =>
                        token.Kind == SpreadsheetNumberFormatTokenKind.BracketedDirective ||
                        token.Kind == SpreadsheetNumberFormatTokenKind.ScalingSeparator ||
                        token.Kind == SpreadsheetNumberFormatTokenKind.TextPlaceholder ||
                        token.Kind == SpreadsheetNumberFormatTokenKind.Literal ||
                        token.Kind == SpreadsheetNumberFormatTokenKind.Other)) unsupported++;
                dataStyles.Add(style.StyleIndex, name);
            }
            target.NumberFormatName = name;
        }
        if (style.Underline || style.Border != null || style.FillGradientUnsupported || style.FillGradientStops.Count > 0 ||
            style.TextRotation.HasValue || style.HorizontalAlignment != null || style.VerticalAlignment != null) unsupported++;
    }

    private static int ApplyOdsStyle(
        ExcelCell target,
        OdsCellRun style,
        IReadOnlyDictionary<string, OdsDataStyle> dataStyles,
        out bool unsupportedDataStyleFormat) {
        int unsupported = 0;
        unsupportedDataStyleFormat = false;
        if (style.Bold == true) target.SetBold();
        if (style.Italic == true) target.SetItalic();
        if (style.FontSize.HasValue) {
            if (style.FontSize.Value.TryToPoints(out double points)) target.SetFontSize(points);
            else unsupported++;
        }
        if (!string.IsNullOrWhiteSpace(style.FontFamily)) target.SetFontName(style.FontFamily!);
        if (style.Color.HasValue) target.SetFontColor(style.Color.Value.ToString().TrimStart('#'));
        if (style.BackgroundColor.HasValue) target.SetFillColor(style.BackgroundColor.Value.ToString().TrimStart('#'));
        if (style.NumberFormatName != null && dataStyles.TryGetValue(style.NumberFormatName, out OdsDataStyle? dataStyle)) {
            if (dataStyle.TryGetExcelNumberFormatCode(out string formatCode)) target.SetNumberFormat(formatCode);
            else unsupportedDataStyleFormat = true;
        }
        return unsupported;
    }

    private static bool IsNumeric(object value) {
        TypeCode code = Type.GetTypeCode(value.GetType());
        return code >= TypeCode.SByte && code <= TypeCode.Decimal;
    }

    private static string ValueText(object? value) => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
    private static double ExcelWidthToPoints(double width) => Math.Max(0D, (width * 7D + 5D) * 72D / 96D);
    private static double PointsToExcelWidth(double points) => Math.Max(0D, Math.Min(255D, (points * 96D / 72D - 5D) / 7D));

    private static long SaturatingAdd(long left, long right) => right > long.MaxValue - left ? long.MaxValue : left + right;

    private static bool IsExternalOdfHref(string href) =>
        !string.IsNullOrWhiteSpace(href) && !href.StartsWith("#", StringComparison.Ordinal)
        && (href.StartsWith("//", StringComparison.Ordinal) || Uri.TryCreate(href, UriKind.Absolute, out _));

    private static void AddConverted(OdfConversionReport report, string feature, int count) {
        if (count > 0) report.Add(feature, OdfConversionMappingStatus.Converted, count);
    }

    private static void AddUnsupported(OdfConversionReport report, string feature, int count, string? message) {
        if (count > 0) report.Add(feature, OdfConversionMappingStatus.Unsupported, count, message);
    }

    private static void AddUnmappedOdfFindings(OdfFeatureReport features, OdfConversionReport report,
        int formulas, int validations, int hyperlinks, int annotations, int namedRanges) {
        foreach (OdfFeatureDiagnostic diagnostic in features.Diagnostics) {
            report.Add("source-inspection", OdfConversionMappingStatus.Unsupported, 1,
                diagnostic.Code + " in " + diagnostic.PartPath + ": " + diagnostic.Message);
        }
        int remainingFormulas = formulas, remainingValidations = validations, remainingHyperlinks = hyperlinks;
        int remainingAnnotations = annotations, remainingNamedRanges = namedRanges;
        foreach (OdfFeatureFinding finding in features.Findings) {
            int handled = 0;
            if (finding.Name == "spreadsheet-formulas") handled = Consume(ref remainingFormulas, finding.Count);
            else if (finding.Name == "spreadsheet-validations") handled = Consume(ref remainingValidations, finding.Count);
            else if (finding.Name == "external-links") handled = Consume(ref remainingHyperlinks, finding.Count);
            else if (finding.Name == "annotations") handled = Consume(ref remainingAnnotations, finding.Count);
            else if (finding.Name == "spreadsheet-named-ranges") handled = Consume(ref remainingNamedRanges, finding.Count);
            int remaining = Math.Max(0, finding.Count - handled);
            if (remaining > 0) report.Add("source-" + finding.Name, OdfConversionMappingStatus.Unsupported, remaining,
                "The source feature is not represented by the XLSX conversion surface.");
        }
    }

    private static int Consume(ref int available, int requested) {
        int consumed = Math.Min(available, requested);
        available -= consumed;
        return consumed;
    }

    private static ExcelDocument Normalize(ExcelDocument document) {
        using var stream = new MemoryStream();
        document.Save(stream);
        document.Dispose();
        using var detachedInput = new MemoryStream(stream.ToArray(), writable: false);
        return ExcelDocument.Load(detachedInput);
    }
}
using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;
using OfficeIMO.Spreadsheet;
using System.Globalization;
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
        target.Metadata.Creator = snapshot.Author;
        target.Metadata.Subject = snapshot.Subject;
        NamedRangeConversionPlan namedRangePlan = BuildNamedRangeConversionPlan(snapshot.NamedRanges);

        int cells = 0, formulas = 0, formulaTranslationFailures = 0, styles = 0, hyperlinks = 0, unsupportedHyperlinks = 0, hyperlinkTooltips = 0, comments = 0, richComments = 0, threadedComments = 0, merges = 0;
        int rows = 0, columns = 0, convertedValidations = 0, skippedValidations = 0, overlappingValidationAssignments = 0;
        int tables = 0, filters = 0, unsupportedStyles = 0, skippedStyles = 0;
        long materializedCells = 0, skippedCells = 0, skippedRows = 0, skippedColumns = 0, skippedMerges = 0;
        bool truncated = false;
        var dataStyles = new Dictionary<uint, string>();
        int worksheetOrdinal = 0;
        foreach (ExcelWorksheetSnapshot worksheet in snapshot.Worksheets) {
            worksheetOrdinal++;
            OdsSheet sheet = target.AddSheet(worksheet.Name);
            var materializedCoordinates = new HashSet<(int Row, int Column)>();
            var validationAssignments = new Dictionary<(int Row, int Column), string>();
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
                    string rewrittenFormula = namedRangePlan.RewriteFormula(cell.Formula!, worksheet.Name);
                    var translation = SpreadsheetAddressConverter.ExcelFormulaToOpenFormula(rewrittenFormula);
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
                    string? href = null;
                    if (cell.Hyperlink.IsExternal) {
                        href = cell.Hyperlink.Target;
                    } else if (namedRangePlan.TryResolveHyperlinkName(cell.Hyperlink.Target, worksheet.Name,
                                   out string outputName)) {
                        href = "#" + outputName;
                    } else {
                        string address = SpreadsheetAddressConverter.ExcelRangeToOpenAddress(cell.Hyperlink.Target);
                        if (address.Length > 0) href = "#" + address;
                    }
                    if (href != null) {
                        converted.SetHyperlink(ValueText(cell.Value), href);
                        hyperlinks++;
                    } else {
                        unsupportedHyperlinks++;
                    }
                    if (!string.IsNullOrWhiteSpace(cell.Hyperlink.Tooltip)) hyperlinkTooltips++;
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
                    if (cell.Comment.RichTextRuns.Any(HasRichTextFormatting)) richComments++;
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
                            if (validationAssignments.TryGetValue(coordinate, out string? previousValidation)
                                && !string.Equals(previousValidation, validationName, StringComparison.Ordinal)) {
                                overlappingValidationAssignments++;
                            }
                            validationAssignments[coordinate] = validationName;
                            sheet.Cell(row - 1L, column - 1L).ValidationName = validationName;
                            assigned = true;
                        }
                        if (validationLimitReached) break;
                    }
                }
                if (assigned) {
                    OdsValidation convertedValidation = target.AddValidation(validationName, condition!, validation.AllowBlank);
                    if (string.Equals(validation.Type, "list", StringComparison.OrdinalIgnoreCase)) {
                        convertedValidation.DisplayList = validation.SuppressDropDown
                            ? OdsValidationDisplayList.None
                            : OdsValidationDisplayList.Unsorted;
                    }
                    if (validation.PromptTitle != null || validation.Prompt != null) {
                        convertedValidation.SetHelpMessage(validation.PromptTitle, validation.Prompt, validation.ShowInputMessage);
                    } else if (validation.ShowInputMessage) {
                        convertedValidation.EnsureHelpMessage();
                    }
                    OdsValidationMessageType messageType = ParseOdsValidationMessageType(validation.ErrorStyle);
                    if (validation.ErrorTitle != null || validation.Error != null) {
                        convertedValidation.SetErrorMessage(
                            validation.ErrorTitle,
                            validation.Error,
                            messageType,
                            validation.ShowErrorMessage);
                    } else if (validation.ShowErrorMessage || messageType != OdsValidationMessageType.Stop) {
                        convertedValidation.EnsureErrorMessage(messageType, validation.ShowErrorMessage);
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

        foreach (NamedRangeConversionEntry named in namedRangePlan.Entries) {
            target.AddNamedRange(named.OutputName, named.Address);
        }
        if (snapshot.ActiveWorksheetIndex.GetValueOrDefault() > 0) {
            report.Add("worksheet-views", OdfConversionMappingStatus.Unsupported, 1,
                "The active Excel worksheet is not represented by the current ODS typed surface; ODS consumers will select their default sheet.");
        }
        int namedRanges = namedRangePlan.Entries.Count;
        int builtInNames = namedRangePlan.BuiltInCount;
        int unsupportedNamedExpressions = namedRangePlan.UnsupportedExpressionCount;
        int disambiguatedNames = namedRangePlan.DisambiguatedCount;

        AddConverted(report, "worksheets", snapshot.Worksheets.Count);
        AddConverted(report, "cells", cells);
        AddConverted(report, "rows", rows);
        if (columns > 0) report.Add("column-layout", OdfConversionMappingStatus.Approximated, columns,
            "Excel character-unit column widths are converted to approximate physical widths.");
        AddConverted(report, "merges", merges);
        AddConverted(report, "hyperlinks", hyperlinks);
        AddUnsupported(report, "hyperlinks", unsupportedHyperlinks,
            "Internal hyperlink targets that are neither cell addresses nor transferred named ranges were omitted.");
        AddUnsupported(report, "hyperlink-tooltips", hyperlinkTooltips,
            "Excel hyperlink ScreenTips have no equivalent in the current ODS hyperlink model.");
        AddConverted(report, "named-ranges", namedRanges);
        if (disambiguatedNames > 0) report.Add("sheet-local-named-ranges", OdfConversionMappingStatus.Approximated, disambiguatedNames,
            "Excel names that collide after ODS workbook-scope projection were made unique, and affected formulas and internal hyperlinks were rewritten to the converted names.");
        if (formulas > 0) report.Add("formulas", OdfConversionMappingStatus.Approximated, formulas,
            "Formula syntax is parsed and translated to OpenFormula; cached values are retained.");
        if (formulaTranslationFailures > 0) report.Add("formulas", OdfConversionMappingStatus.Skipped, formulaTranslationFailures,
            "Formula syntax without a safe OpenFormula representation was omitted; the cached cell value was retained.");
        if (styles > 0) report.Add("cell-styles", OdfConversionMappingStatus.Approximated, styles,
            "Bold, italic, font, foreground, fill, and common number formats are mapped; other Excel style details are omitted.");
        if (skippedStyles > 0) report.Add("cell-styles", OdfConversionMappingStatus.Skipped, skippedStyles,
            "Cell styles were omitted because IncludeBasicStyles is disabled.");
        if (unsupportedStyles > 0) report.Add("cell-format-details", OdfConversionMappingStatus.Unsupported, unsupportedStyles);
        AddConverted(report, "comments", comments - richComments);
        if (richComments > 0) report.Add("comments", OdfConversionMappingStatus.Approximated, richComments,
            "Legacy Excel comment text and authors are retained, but rich-text run formatting is flattened in ODS annotations.");
        if (threadedComments > 0) report.Add("threaded-comments", OdfConversionMappingStatus.Approximated, threadedComments,
            "Each cell thread was flattened into one schema-valid ODS annotation transcript that retains comment bodies and available author, timestamp, identity, parent, and resolved-state metadata.");
        AddConverted(report, "validations", convertedValidations);
        if (skippedValidations > 0) report.Add("validations", OdfConversionMappingStatus.Unsupported, skippedValidations,
            "Unsupported validation rules and ranges, plus assignments clipped by configured conversion limits, were not mapped completely.");
        AddUnsupported(report, "validation-overlaps", overlappingValidationAssignments,
            "ODF cells can reference only one validation rule; where Excel rules overlap, the later workbook rule was retained for that cell.");
        AddUnsupported(report, "structured-tables", tables, "Table cells remain; Excel table semantics and styles are not translated.");
        AddUnsupported(report, "filters", filters, "Filter state is not translated.");
        AddUnsupported(report, "built-in-names", builtInNames, "Excel print-area and print-title names are not translated.");
        AddUnsupported(report, "named-expressions", unsupportedNamedExpressions,
            "Excel defined names that contain constants or formulas instead of representable A1 ranges are not translated to ODS.");
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
        target.BuiltinDocumentProperties.Creator = source.Metadata.Creator;
        target.BuiltinDocumentProperties.Subject = source.Metadata.Subject;
        var dataStyles = source.DataStyles.GroupBy(style => style.Name, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        var sourceValidations = source.Validations
            .GroupBy(validation => validation.Name, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        OdsNamedRangeConversionPlan namedRangePlan = BuildOdsNamedRangeConversionPlan(source.NamedRanges);
        CultureInfo textCaseCulture = OdfTextCultureResolver.Resolve(source.Metadata.Language);

        long expandedCells = 0;
        int cells = 0, formulas = 0, formulaTranslationFailures = 0, styles = 0, hyperlinks = 0, externalHyperlinks = 0, comments = 0, combinedComments = 0, metadataTranscriptComments = 0, merges = 0, rowLayouts = 0, columnLayouts = 0;
        int invalidValues = 0, normalizedDateTimeOffsets = 0, validations = 0, convertedValidations = 0, unsupportedValidationAssignments = 0, sortedValidationLists = 0, unsupportedHyperlinks = 0, unsupportedMeasurements = 0, unsupportedDataStyleFormats = 0, skippedStyles = 0, renamedSheets = 0, worksheetCount = 0;
        int approximatedFontFamilyLists = 0, unsupportedFontFamilies = 0;
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
                            ExcelValueProjectionStatus valueStatus = SetExcelValue(converted, cellRun.Value);
                            if (valueStatus == ExcelValueProjectionStatus.Invalid) invalidValues++;
                            else if (valueStatus == ExcelValueProjectionStatus.TimeZoneNormalized) normalizedDateTimeOffsets++;
                            if (!string.IsNullOrWhiteSpace(cellRun.Formula)) {
                                var translation = SpreadsheetAddressConverter.OpenFormulaToExcel(cellRun.Formula!);
                                if (translation.IsSuccessful) {
                                    converted.SetFormula(namedRangePlan.RewriteFormula(translation.Formula));
                                    formulas++;
                                } else {
                                    formulaTranslationFailures++;
                                }
                            }
                            if (!string.IsNullOrWhiteSpace(cellRun.HyperlinkHref)) {
                                string href = cellRun.HyperlinkHref!;
                                bool convertedHyperlink = false;
                                if (OdfUriReference.TryDecodeFragment(href, out string fragment)) {
                                    string location = SpreadsheetAddressConverter.OpenAddressToExcel(fragment);
                                    if (location.Length == 0 && namedRangePlan.TryResolveName(fragment, out string outputName)) {
                                        location = outputName;
                                    }
                                    if (location.Length > 0) {
                                        sheet.SetInternalLink(excelRow, excelColumn, location, cellRun.Text, style: true);
                                        convertedHyperlink = true;
                                    } else {
                                        unsupportedHyperlinks++;
                                    }
                                } else if (!href.StartsWith("#", StringComparison.Ordinal)) {
                                    sheet.SetHyperlink(excelRow, excelColumn, href, cellRun.Text, style: true);
                                    convertedHyperlink = true;
                                } else {
                                    unsupportedHyperlinks++;
                                }
                                if (cellRun.Value.Kind != OdsCellValueKind.Empty) _ = SetExcelValue(converted, cellRun.Value);
                                if (!string.IsNullOrWhiteSpace(cellRun.Formula)) {
                                    var translation = SpreadsheetAddressConverter.OpenFormulaToExcel(cellRun.Formula!);
                                    if (translation.IsSuccessful) {
                                        converted.SetFormula(namedRangePlan.RewriteFormula(translation.Formula));
                                    }
                                }
                                if (convertedHyperlink) {
                                    hyperlinks++;
                                    if (IsExternalOdfHref(href)) externalHyperlinks++;
                                }
                            }
                            if (effective.IncludeBasicStyles && cellRun.StyleName != null) {
                                unsupportedMeasurements += ApplyOdsStyle(converted, cellRun, dataStyles,
                                    out bool unsupportedDataStyleFormat, ref approximatedFontFamilyLists,
                                    ref unsupportedFontFamilies, textCaseCulture);
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
                if (validation.ParsedCondition?.ValueKind == OdsValidationValueKind.List
                    && validation.DisplayList == OdsValidationDisplayList.SortAscending) sortedValidationLists++;
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

        foreach (NamedRangeConversionEntry named in namedRangePlan.Entries) {
            target.SetNamedRange(named.OutputName, named.Address, save: false,
                validationMode: ExcelDefinedNameValidationMode.Strict);
        }
        int namedRanges = namedRangePlan.Entries.Count;

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
        if (namedRangePlan.RenamedCount > 0) report.Add("named-range-names",
            OdfConversionMappingStatus.Approximated, namedRangePlan.RenamedCount,
            "ODF names that are not legal Excel defined names were sanitized, and formulas and internal hyperlinks were rewritten to the emitted names.");
        if (formulas > 0) report.Add("formulas", OdfConversionMappingStatus.Approximated, formulas,
            "OpenFormula syntax is parsed and translated to Excel A1 syntax; cached ODS values remain independently represented.");
        if (formulaTranslationFailures > 0) report.Add("formulas", OdfConversionMappingStatus.Skipped, formulaTranslationFailures,
            "Formula syntax without a safe Excel A1 representation was omitted; the cached cell value was retained.");
        if (styles > 0) report.Add("cell-styles", OdfConversionMappingStatus.Approximated, styles,
            "Basic font, fill, and data-style categories are mapped.");
        if (approximatedFontFamilyLists > 0) report.Add("font-family-fallbacks", OdfConversionMappingStatus.Approximated,
            approximatedFontFamilyLists, "Excel cell styles retain the first ODF font family but cannot retain the authored fallback list.");
        AddUnsupported(report, "font-families", unsupportedFontFamilies,
            "Malformed ODF font-family syntax was omitted instead of being emitted as an invalid Excel typeface name.");
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
        AddUnsupported(report, "validation-display-lists", sortedValidationLists,
            "Excel preserves the authored validation-list order but cannot request ODF's ascending display order.");
        AddUnsupported(report, "invalid-values", invalidValues, "Invalid typed lexemes were transferred as display text.");
        AddUnsupported(report, "date-time-offsets", normalizedDateTimeOffsets,
            "Offset-bearing ODF date/time values were normalized to their UTC instant before Excel serial storage; Excel cannot retain the authored offset.");
        AddUnsupported(report, "relative-measurements", unsupportedMeasurements,
            "Relative or unsupported ODF row, column, or text measurements could not be projected to fixed Excel sizes and were omitted.");
        AddUnsupported(report, "cell-format-details", unsupportedDataStyleFormats,
            $"ODF data styles with locale-sensitive or unsupported components, or that exceed Excel's {OdsDataStyle.MaximumExcelNumberFormatCodeLength}-character custom number-format limit, were omitted.");
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

    private static string ValueText(object? value) => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
    private static double ExcelWidthToPoints(double width) => Math.Max(0D, (width * 7D + 5D) * 72D / 96D);
    private static double PointsToExcelWidth(double points) => Math.Max(0D, Math.Min(255D, (points * 96D / 72D - 5D) / 7D));

    private static long SaturatingAdd(long left, long right) => right > long.MaxValue - left ? long.MaxValue : left + right;

    private static bool IsExternalOdfHref(string href) =>
        !string.IsNullOrWhiteSpace(href) && !href.StartsWith("#", StringComparison.Ordinal)
        && (href.StartsWith("//", StringComparison.Ordinal) || Uri.TryCreate(href, UriKind.Absolute, out _));

    private static ExcelDocument Normalize(ExcelDocument document) {
        using var stream = new MemoryStream();
        document.Save(stream);
        document.Dispose();
        using var detachedInput = new MemoryStream(stream.ToArray(), writable: false);
        return ExcelDocument.Load(detachedInput);
    }
}

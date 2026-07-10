using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;

namespace OfficeIMO.Excel.OpenDocument;

/// <summary>Explicit conversions between OfficeIMO Excel and native OpenDocument spreadsheet models.</summary>
public static class ExcelOpenDocumentConversionExtensions {
    /// <summary>Converts an Excel workbook to an in-memory ODS document and reports every lossy mapping.</summary>
    public static OdfConversionResult<OdsDocument> ToOpenDocument(this ExcelDocument source,
        ExcelOpenDocumentConversionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        ExcelOpenDocumentConversionOptions effective = options ?? new ExcelOpenDocumentConversionOptions();
        effective.Validate();
        ExcelWorkbookSnapshot snapshot = source.CreateInspectionSnapshot();
        OdsDocument target = OdsDocument.Create();
        var report = new OdfConversionReport("XLSX", "ODS");
        target.Metadata.Title = snapshot.Title;

        int cells = 0, formulas = 0, styles = 0, hyperlinks = 0, comments = 0, merges = 0;
        int rows = 0, columns = 0, validations = 0, tables = 0, filters = 0, unsupportedStyles = 0;
        var dataStyles = new Dictionary<uint, string>();
        foreach (ExcelWorksheetSnapshot worksheet in snapshot.Worksheets) {
            OdsSheet sheet = target.AddSheet(worksheet.Name);
            sheet.Hidden = worksheet.Hidden;
            foreach (ExcelColumnSnapshot column in worksheet.Columns) {
                int last = Math.Min(column.EndIndex, effective.MaximumColumns);
                for (int index = Math.Max(1, column.StartIndex); index <= last; index++) {
                    OdsColumn converted = sheet.Column(index - 1L);
                    converted.Hidden = column.Hidden;
                    if (column.Width.HasValue) converted.Width = OdfLength.Points(ExcelWidthToPoints(column.Width.Value));
                    columns++;
                }
            }
            foreach (ExcelRowSnapshot row in worksheet.Rows) {
                if (row.Index < 1 || row.Index > effective.MaximumRows) continue;
                OdsRow converted = sheet.Row(row.Index - 1L);
                converted.Hidden = row.Hidden;
                if (row.Height.HasValue) converted.Height = OdfLength.Points(row.Height.Value);
                rows++;
            }

            foreach (ExcelCellSnapshot cell in worksheet.Cells) {
                if (cell.Row < 1 || cell.Column < 1 || cell.Row > effective.MaximumRows || cell.Column > effective.MaximumColumns) continue;
                OdsCell converted = sheet.Cell(cell.Row - 1L, cell.Column - 1L);
                if (!string.IsNullOrWhiteSpace(cell.Formula)) {
                    converted.Formula = SpreadsheetAddressConverter.ExcelFormulaToOpenFormula(cell.Formula!);
                    formulas++;
                }
                bool exactValue = SetOdsValue(converted, cell.Value);
                if (!exactValue) unsupportedStyles++;
                if (cell.Hyperlink != null && !string.IsNullOrWhiteSpace(cell.Hyperlink.Target)) {
                    converted.SetHyperlink(ValueText(cell.Value), cell.Hyperlink.Target);
                    hyperlinks++;
                }
                if (effective.IncludeBasicStyles && cell.Style != null) {
                    ApplyExcelStyle(target, converted, cell.Style, dataStyles, ref unsupportedStyles);
                    styles++;
                }
                if (cell.Comment != null) comments++;
                if (cell.ThreadedComment != null) comments++;
                cells++;
            }

            foreach (ExcelMergedRangeSnapshot merged in worksheet.MergedRanges) {
                if (merged.StartRow < 1 || merged.StartColumn < 1 || merged.StartRow > effective.MaximumRows || merged.StartColumn > effective.MaximumColumns) continue;
                long rowSpan = Math.Min(merged.EndRow, effective.MaximumRows) - merged.StartRow + 1L;
                long columnSpan = Math.Min(merged.EndColumn, effective.MaximumColumns) - merged.StartColumn + 1L;
                sheet.Merge(merged.StartRow - 1L, merged.StartColumn - 1L, rowSpan, columnSpan);
                merges++;
            }
            validations += worksheet.Validations.Count;
            tables += worksheet.Tables.Count;
            if (worksheet.AutoFilter != null) filters++;
            if (worksheet.FrozenRowCount > 0 || worksheet.FrozenColumnCount > 0 || worksheet.RightToLeft || !worksheet.ShowGridlines) {
                report.Add("worksheet-views", OdfConversionMappingStatus.Unsupported, 1,
                    "Frozen panes and Excel-specific worksheet view settings are not represented by the current ODS typed surface.");
            }
            if (worksheet.Protection != null) report.Add("worksheet-protection", OdfConversionMappingStatus.Unsupported, 1);
        }

        int namedRanges = 0, builtInNames = 0;
        foreach (ExcelNamedRangeSnapshot named in snapshot.NamedRanges) {
            if (named.IsBuiltIn) { builtInNames++; continue; }
            string address = SpreadsheetAddressConverter.ExcelRangeToOpenAddress(named.ReferenceA1, named.SheetName);
            if (address.Length == 0) continue;
            target.AddNamedRange(named.Name, address);
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
        if (formulas > 0) report.Add("formulas", OdfConversionMappingStatus.Approximated, formulas,
            "Formula text and cached values are retained; local A1 references are translated to an OpenFormula subset.");
        if (styles > 0) report.Add("cell-styles", OdfConversionMappingStatus.Approximated, styles,
            "Bold, italic, font, foreground, fill, and common number formats are mapped; other Excel style details are omitted.");
        if (unsupportedStyles > 0) report.Add("cell-format-details", OdfConversionMappingStatus.Unsupported, unsupportedStyles);
        AddUnsupported(report, "comments", comments, "ODS annotations are not exposed by the current native spreadsheet model.");
        AddUnsupported(report, "validations", validations, "Cell values remain, but Excel validation rules are not translated yet.");
        AddUnsupported(report, "structured-tables", tables, "Table cells remain; Excel table semantics and styles are not translated.");
        AddUnsupported(report, "filters", filters, "Filter state is not translated.");
        AddUnsupported(report, "built-in-names", builtInNames, "Excel print-area and print-title names are not translated.");
        AddUnsupported(report, "slicers", snapshot.SlicerPartCount, null);
        AddUnsupported(report, "timelines", snapshot.TimelinePartCount, null);
        AddUnsupported(report, "connections", snapshot.ConnectionPartCount, null);
        AddUnsupported(report, "query-tables", snapshot.QueryTablePartCount, null);
        return new OdfConversionResult<OdsDocument>(target, report);
    }

    /// <summary>Converts an ODS document to an in-memory Excel workbook and reports every lossy mapping.</summary>
    public static OdfConversionResult<ExcelDocument> ToExcelDocument(this OdsDocument source,
        ExcelOpenDocumentConversionOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        ExcelOpenDocumentConversionOptions effective = options ?? new ExcelOpenDocumentConversionOptions();
        effective.Validate();
        ExcelDocument target = ExcelDocument.Create(new MemoryStream(), autoSave: false);
        var report = new OdfConversionReport("ODS", "XLSX");
        target.BuiltinDocumentProperties.Title = source.Metadata.Title;
        var dataStyles = source.DataStyles.ToDictionary(style => style.Name, style => style.Kind, StringComparer.Ordinal);

        long expandedCells = 0;
        int cells = 0, formulas = 0, styles = 0, hyperlinks = 0, merges = 0, rowLayouts = 0, columnLayouts = 0;
        int invalidValues = 0, validations = 0;
        bool truncated = false;
        foreach (OdsSheet odsSheet in source.Sheets) {
            ExcelSheet sheet = target.AddWorkSheet(odsSheet.Name);
            sheet.SetHidden(odsSheet.Hidden);

            foreach (OdsColumnRun columnRun in odsSheet.ColumnRuns) {
                long lastExclusive = Math.Min(checked(columnRun.StartColumn + columnRun.RepeatCount), effective.MaximumColumns);
                for (long column = columnRun.StartColumn; column < lastExclusive; column++) {
                    if (!columnRun.Hidden && !columnRun.Width.HasValue) continue;
                    int excelColumn = checked((int)column + 1);
                    if (columnRun.Hidden) sheet.SetColumnHidden(excelColumn, true);
                    if (columnRun.Width.HasValue) sheet.SetColumnWidth(excelColumn, PointsToExcelWidth(columnRun.Width.Value.ToPoints()));
                    columnLayouts++;
                }
                if (checked(columnRun.StartColumn + columnRun.RepeatCount) > effective.MaximumColumns) truncated = true;
            }

            foreach (OdsRowRun rowRun in odsSheet.RowRuns) {
                long rowEnd = checked(rowRun.StartRow + rowRun.RepeatCount);
                long lastRowExclusive = Math.Min(rowEnd, effective.MaximumRows);
                if (rowEnd > effective.MaximumRows) truncated = true;
                for (long row = rowRun.StartRow; row < lastRowExclusive; row++) {
                    int excelRow = checked((int)row + 1);
                    if (rowRun.Hidden) { sheet.SetRowHidden(excelRow, true); rowLayouts++; }
                    if (rowRun.Height.HasValue) { sheet.SetRowHeight(excelRow, rowRun.Height.Value.ToPoints()); rowLayouts++; }

                    foreach (OdsCellRun cellRun in rowRun.CellRuns) {
                        long columnEnd = checked(cellRun.StartColumn + cellRun.RepeatCount);
                        long lastColumnExclusive = Math.Min(columnEnd, effective.MaximumColumns);
                        if (columnEnd > effective.MaximumColumns) truncated = true;
                        if (cellRun.IsCovered || !IsSignificant(cellRun)) continue;
                        for (long column = cellRun.StartColumn; column < lastColumnExclusive; column++) {
                            if (expandedCells >= effective.MaximumExpandedCells) { truncated = true; break; }
                            expandedCells++;
                            int excelColumn = checked((int)column + 1);
                            ExcelCell converted = sheet.CellAt(excelRow, excelColumn);
                            if (!SetExcelValue(converted, cellRun.Value)) invalidValues++;
                            if (!string.IsNullOrWhiteSpace(cellRun.Formula)) {
                                converted.SetFormula(SpreadsheetAddressConverter.OpenFormulaToExcel(cellRun.Formula!));
                                formulas++;
                            }
                            if (!string.IsNullOrWhiteSpace(cellRun.HyperlinkHref)) {
                                sheet.SetHyperlink(excelRow, excelColumn, cellRun.HyperlinkHref!, cellRun.Text, style: true);
                                hyperlinks++;
                            }
                            if (effective.IncludeBasicStyles && cellRun.StyleName != null) {
                                ApplyOdsStyle(converted, cellRun, dataStyles);
                                styles++;
                            }
                            if (cellRun.ValidationName != null) validations++;
                            cells++;

                            if (cellRun.RowSpan > 1 || cellRun.ColumnSpan > 1) {
                                long mergeLastRow = row + cellRun.RowSpan;
                                long mergeLastColumn = column + cellRun.ColumnSpan;
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
            if (expandedCells >= effective.MaximumExpandedCells) break;
        }

        if (target.Sheets.Count == 0) target.AddWorkSheet("Sheet1");
        OdsSheet? activeSource = source.Sheets.FirstOrDefault(sheet => !sheet.Hidden);
        if (activeSource != null) target.SetActiveWorksheet(activeSource.Name);

        int namedRanges = 0;
        foreach (OdsNamedRange named in source.NamedRanges) {
            string reference = SpreadsheetAddressConverter.OpenAddressToExcel(named.CellRangeAddress);
            if (reference.Length == 0) continue;
            target.SetNamedRange(named.Name, reference, save: false);
            namedRanges++;
        }

        AddConverted(report, "worksheets", source.Sheets.Count);
        AddConverted(report, "cells", cells);
        AddConverted(report, "row-layout", rowLayouts);
        if (columnLayouts > 0) report.Add("column-layout", OdfConversionMappingStatus.Approximated, columnLayouts,
            "Physical ODF column widths are converted to approximate Excel character widths.");
        AddConverted(report, "merges", merges);
        AddConverted(report, "hyperlinks", hyperlinks);
        AddConverted(report, "named-ranges", namedRanges);
        if (formulas > 0) report.Add("formulas", OdfConversionMappingStatus.Approximated, formulas,
            "OpenFormula text is translated to an Excel formula subset; cached ODS values remain available only when independently represented.");
        if (styles > 0) report.Add("cell-styles", OdfConversionMappingStatus.Approximated, styles,
            "Basic font, fill, and data-style categories are mapped.");
        AddUnsupported(report, "validations", validations, "ODF validation conditions are retained in the source but are not translated to Excel rules.");
        AddUnsupported(report, "invalid-values", invalidValues, "Invalid typed lexemes were transferred as display text.");
        if (truncated) report.Add("expansion-limits", OdfConversionMappingStatus.Skipped, 1,
            "Content outside the configured row, column, or expanded-cell limits was not materialized.");
        target = Normalize(target);
        return new OdfConversionResult<ExcelDocument>(target, report);
    }

    private static bool IsSignificant(OdsCellRun cell) => cell.Value.Kind != OdsCellValueKind.Empty ||
        cell.Formula != null || cell.StyleName != null || cell.ValidationName != null || cell.HyperlinkHref != null ||
        cell.RowSpan > 1 || cell.ColumnSpan > 1;

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
                if (style.IsDateLike) document.AddDateStyle(name);
                else if (style.NumberFormatCode!.IndexOf('%') >= 0) document.AddPercentageStyle(name, CountDecimalPlaces(style.NumberFormatCode));
                else if (TryCurrencySymbol(style.NumberFormatCode, out string symbol)) document.AddCurrencyStyle(name, symbol, CountDecimalPlaces(style.NumberFormatCode));
                else document.AddNumberStyle(name, CountDecimalPlaces(style.NumberFormatCode));
                dataStyles.Add(style.StyleIndex, name);
            }
            target.NumberFormatName = name;
        }
        if (style.Underline || style.Border != null || style.FillGradientUnsupported || style.FillGradientStops.Count > 0 ||
            style.TextRotation.HasValue || style.HorizontalAlignment != null || style.VerticalAlignment != null) unsupported++;
    }

    private static void ApplyOdsStyle(ExcelCell target, OdsCellRun style, IReadOnlyDictionary<string, OdsDataStyleKind> dataStyles) {
        if (style.Bold == true) target.SetBold();
        if (style.Italic == true) target.SetItalic();
        if (style.FontSize.HasValue) target.SetFontSize(style.FontSize.Value.ToPoints());
        if (!string.IsNullOrWhiteSpace(style.FontFamily)) target.SetFontName(style.FontFamily!);
        if (style.Color.HasValue) target.SetFontColor(style.Color.Value.ToString().TrimStart('#'));
        if (style.BackgroundColor.HasValue) target.SetFillColor(style.BackgroundColor.Value.ToString().TrimStart('#'));
        if (style.NumberFormatName != null && dataStyles.TryGetValue(style.NumberFormatName, out OdsDataStyleKind kind)) {
            switch (kind) {
                case OdsDataStyleKind.Percentage: target.SetNumberFormat("0.00%"); break;
                case OdsDataStyleKind.Currency: target.SetNumberFormat("#,##0.00"); break;
                case OdsDataStyleKind.Date: target.SetNumberFormat("yyyy-mm-dd"); break;
                case OdsDataStyleKind.Time: target.SetNumberFormat("hh:mm:ss"); break;
                default: target.SetNumberFormat("#,##0.00"); break;
            }
        }
    }

    private static bool IsNumeric(object value) {
        TypeCode code = Type.GetTypeCode(value.GetType());
        return code >= TypeCode.SByte && code <= TypeCode.Decimal;
    }

    private static string ValueText(object? value) => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
    private static double ExcelWidthToPoints(double width) => Math.Max(0D, (width * 7D + 5D) * 72D / 96D);
    private static double PointsToExcelWidth(double points) => Math.Max(0D, Math.Min(255D, (points * 96D / 72D - 5D) / 7D));

    private static int CountDecimalPlaces(string format) {
        int dot = format.IndexOf('.');
        if (dot < 0) return 0;
        int count = 0;
        for (int index = dot + 1; index < format.Length && (format[index] == '0' || format[index] == '#'); index++) count++;
        return Math.Min(10, count);
    }

    private static bool TryCurrencySymbol(string format, out string symbol) {
        foreach (char candidate in new[] { '$', '€', '£', '¥' }) {
            if (format.IndexOf(candidate) >= 0) { symbol = candidate.ToString(); return true; }
        }
        symbol = string.Empty;
        return false;
    }

    private static void AddConverted(OdfConversionReport report, string feature, int count) {
        if (count > 0) report.Add(feature, OdfConversionMappingStatus.Converted, count);
    }

    private static void AddUnsupported(OdfConversionReport report, string feature, int count, string? message) {
        if (count > 0) report.Add(feature, OdfConversionMappingStatus.Unsupported, count, message);
    }

    private static ExcelDocument Normalize(ExcelDocument document) {
        using var stream = new MemoryStream();
        document.Save(stream);
        document.Dispose();
        stream.Position = 0;
        return ExcelDocument.Load(stream, autoSave: false);
    }
}

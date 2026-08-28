using AngleSharp.Dom;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    private static void ImportTableGrid(
        IElement table,
        ExcelSheet sheet,
        HtmlToExcelResult result,
        HtmlToExcelOptions options,
        HtmlImportBudget budget,
        int firstRow,
        int firstColumn,
        HashSet<long>? importedFormulaCells,
        bool useSemanticValues) {
        int maxTableCells = budget.Limits.MaxTableCells;

        var occupiedCells = new HashSet<long>();
        int rowOffset = 0;

        foreach (IElement row in EnumerateDirectTableRows(table)) {
            int rowIndex = firstRow + rowOffset;
            if (rowIndex > A1.MaxRows) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "HTML table rows exceeded the Excel worksheet row limit; remaining rows were skipped.", lossKind: OfficeConversionLossKind.Omission);
                break;
            }

            int columnIndex = firstColumn;
            foreach (IElement cell in row.Children.Where(IsTableCell)) {
                while (occupiedCells.Contains(GetImportCellKey(rowIndex, columnIndex))) {
                    columnIndex++;
                }

                if (columnIndex > A1.MaxColumns) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "HTML table columns exceeded the Excel worksheet column limit in row " + rowIndex.ToString(CultureInfo.InvariantCulture) + "; remaining cells in the row were skipped.", lossKind: OfficeConversionLossKind.Omission);
                    break;
                }

                int cellRow = rowIndex;
                int cellColumn = columnIndex;
                string? semanticReference = cell.GetAttribute("data-officeimo-cell");
                if (TryParseCellReference(semanticReference, out int semanticRow, out int semanticColumn)
                    && semanticRow <= A1.MaxRows
                    && semanticColumn <= A1.MaxColumns) {
                    cellRow = semanticRow;
                    cellColumn = semanticColumn;
                } else if (!string.IsNullOrWhiteSpace(semanticReference)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                        "Cell coordinate '" + semanticReference + "' was outside the Excel worksheet grid and the table position was used instead.", lossKind: OfficeConversionLossKind.Approximation);
                }

                if (occupiedCells.Contains(GetImportCellKey(cellRow, cellColumn))) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TableSpanInvalid,
                        "Cell " + BuildCellReference(cellRow, cellColumn) + " overlapped an earlier HTML table span and was moved to the next available column.", lossKind: OfficeConversionLossKind.Approximation);
                    cellRow = rowIndex;
                    cellColumn = columnIndex;
                }

                int rowSpan = ReadSpan(cell, "rowspan", cellRow, A1.MaxRows, cellRow, cellColumn, result);
                int columnSpan = ReadSpan(cell, "colspan", cellColumn, A1.MaxColumns, cellRow, cellColumn, result);
                long spanArea = (long)rowSpan * columnSpan;
                if (spanArea > maxTableCells - occupiedCells.Count) {
                    if (occupiedCells.Count >= maxTableCells) {
                        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                            "HTML table exceeded the configured MaxTableCells limit; remaining cells were skipped.", lossKind: OfficeConversionLossKind.Omission);
                        return;
                    }

                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "Cell " + BuildCellReference(cellRow, cellColumn) + " contained a span that exceeded the configured MaxTableCells limit; the span was ignored.", lossKind: OfficeConversionLossKind.Approximation);
                    rowSpan = 1;
                    columnSpan = 1;
                }

                if (SpanOverlaps(occupiedCells, cellRow, cellColumn, rowSpan, columnSpan)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TableSpanInvalid,
                        "Cell " + BuildCellReference(cellRow, cellColumn) + " contained an overlapping HTML table span; the span was ignored.", lossKind: OfficeConversionLossKind.Approximation);
                    rowSpan = 1;
                    columnSpan = 1;
                }

                ReserveSpan(occupiedCells, cellRow, cellColumn, rowSpan, columnSpan);

                string text = NormalizeText(cell.TextContent);
                ExcelCell targetCell = sheet.CellAt(cellRow, cellColumn);
                if (!IsSemanticEmptyCell(cell) && (text.Length > 0 || cell.GetAttribute("data-officeimo-value") != null)) {
                    if (SetCellValue(sheet, cellRow, cellColumn, cell, text, result, options, budget, importedFormulaCells, useSemanticValues)) {
                        result.Cells++;
                    }
                }
                ApplyImportedCellTextFormatting(cell, targetCell);

                if (rowSpan > 1 || columnSpan > 1) {
                    sheet.MergeRange(BuildRangeReference(cellRow, cellColumn, cellRow + rowSpan - 1, cellColumn + columnSpan - 1));
                    result.MergedRanges++;
                }

                columnIndex = Math.Max(columnIndex, cellColumn + columnSpan);
            }

            rowOffset++;
        }
    }

    private static void ApplyImportedCellTextFormatting(IElement source, ExcelCell target) {
        IReadOnlyDictionary<string, string> cellCss = ParseInlineStyle(source.GetAttribute("style"));
        ApplyImportedCellStyle(source, target, cellCss);
        IElement[] spans = source.QuerySelectorAll("span").ToArray();
        if (spans.Length > 0) {
            var runs = new List<ExcelRichTextRun>(spans.Length);
            foreach (IElement span in spans) {
                ExcelRichTextRun run = CreateImportedRichTextRun(span, source, cellCss);
                if (run.Text.Length > 0) runs.Add(run);
            }
            if (runs.Count > 0 && string.Equals(string.Concat(runs.Select(run => run.Text)), source.TextContent, StringComparison.Ordinal)) {
                target.SetRichText(runs.ToArray());
                return;
            }
        }
    }

    private static void ApplyImportedCellStyle(IElement source, ExcelCell target, IReadOnlyDictionary<string, string> css) {
        if (IsCssBold(css)) target.SetBold();
        if (IsCssItalic(css)) target.SetItalic();
        ExcelUnderlineStyle? underline = ResolveImportedUnderline(source, css);
        if (underline.HasValue && underline.Value != ExcelUnderlineStyle.None) target.SetUnderline(underline.Value);
        if (HasDecoration(css, "line-through")) target.SetStrikethrough();
        string nativeVerticalAlign = (source.GetAttribute("data-officeimo-excel-vertical-align") ?? string.Empty).Trim();
        if (Enum.TryParse(nativeVerticalAlign, ignoreCase: true, out ExcelVerticalTextAlignment nativeAlignment)
            && Enum.IsDefined(typeof(ExcelVerticalTextAlignment), nativeAlignment)) {
            target.SetVerticalTextAlignment(nativeAlignment);
        } else if (TryGetCss(css, "vertical-align", out string verticalAlign)) {
            if (verticalAlign.Equals("super", StringComparison.OrdinalIgnoreCase)) target.SetSuperscript();
            if (verticalAlign.Equals("sub", StringComparison.OrdinalIgnoreCase)) target.SetSubscript();
            if (verticalAlign.Equals("baseline", StringComparison.OrdinalIgnoreCase)) target.SetBaseline();
        }
        if (TryGetCss(css, "font-family", out string fontFamily)) target.SetFontName(NormalizeCssFontFamily(fontFamily));
        if (TryGetCssPoints(css, "font-size", out double fontSize)) target.SetFontSize(fontSize);
        if (TryGetCss(css, "color", out string color) && TryNormalizeCssHex(color, out string fontColor)) target.SetFontColor(fontColor);
    }

    private static ExcelRichTextRun CreateImportedRichTextRun(
        IElement source,
        IElement inheritedSource,
        IReadOnlyDictionary<string, string> inheritedCss) {
        IReadOnlyDictionary<string, string> css = ParseInlineStyle(source.GetAttribute("style"));
        ExcelUnderlineStyle? underline = ResolveImportedUnderline(source, css)
            ?? ResolveImportedUnderline(inheritedSource, inheritedCss);
        var run = new ExcelRichTextRun(source.TextContent) {
            Bold = HasCss(css, "font-weight") ? IsCssBold(css) : IsCssBold(inheritedCss),
            Italic = HasCss(css, "font-style") ? IsCssItalic(css) : IsCssItalic(inheritedCss),
            Underline = underline.HasValue && underline.Value != ExcelUnderlineStyle.None,
            UnderlineStyle = underline,
            Strikethrough = HasDecorationDeclaration(css)
                ? HasDecoration(css, "line-through")
                : HasDecoration(inheritedCss, "line-through")
        };
        if (!TryGetCss(css, "vertical-align", out string verticalAlign)) {
            TryGetCss(inheritedCss, "vertical-align", out verticalAlign);
        }
        if (!string.IsNullOrEmpty(verticalAlign)) {
            if (verticalAlign.Equals("super", StringComparison.OrdinalIgnoreCase)) run.VerticalTextAlignment = ExcelVerticalTextAlignment.Superscript;
            if (verticalAlign.Equals("sub", StringComparison.OrdinalIgnoreCase)) run.VerticalTextAlignment = ExcelVerticalTextAlignment.Subscript;
            if (verticalAlign.Equals("baseline", StringComparison.OrdinalIgnoreCase)) run.VerticalTextAlignment = ExcelVerticalTextAlignment.Baseline;
        }
        if (!TryGetCss(css, "font-family", out string fontFamily)) TryGetCss(inheritedCss, "font-family", out fontFamily);
        if (!string.IsNullOrEmpty(fontFamily)) run.FontName = NormalizeCssFontFamily(fontFamily);
        if (!TryGetCssPoints(css, "font-size", out double fontSize)) TryGetCssPoints(inheritedCss, "font-size", out fontSize);
        if (fontSize > 0D) run.FontSize = fontSize;
        if (!TryGetCss(css, "color", out string color)) TryGetCss(inheritedCss, "color", out color);
        if (!string.IsNullOrEmpty(color) && TryNormalizeCssHex(color, out string fontColor)) run.FontColor = fontColor;
        return run;
    }

    private static ExcelUnderlineStyle? ResolveImportedUnderline(IElement source, IReadOnlyDictionary<string, string> css) {
        string exact = (source.GetAttribute("data-officeimo-excel-underline") ?? string.Empty).Trim();
        if (Enum.TryParse(exact, ignoreCase: true, out ExcelUnderlineStyle native)
            && Enum.IsDefined(typeof(ExcelUnderlineStyle), native)) return native;
        if (!HasDecoration(css, "underline")) return null;
        return TryGetCss(css, "text-decoration-style", out string style)
            && style.Equals("double", StringComparison.OrdinalIgnoreCase)
            ? ExcelUnderlineStyle.Double
            : ExcelUnderlineStyle.Single;
    }

    private static IReadOnlyDictionary<string, string> ParseInlineStyle(string? value) {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (string declaration in (value ?? string.Empty).Split(';')) {
            int separator = declaration.IndexOf(':');
            if (separator <= 0) continue;
            string name = declaration.Substring(0, separator).Trim();
            string content = declaration.Substring(separator + 1).Trim();
            if (name.Length > 0 && content.Length > 0) result[name] = content;
        }
        return result;
    }

    private static bool TryGetCss(IReadOnlyDictionary<string, string> css, string name, out string value) =>
        css.TryGetValue(name, out value!);

    private static bool HasCss(IReadOnlyDictionary<string, string> css, string name) => css.ContainsKey(name);

    private static bool HasDecorationDeclaration(IReadOnlyDictionary<string, string> css) =>
        HasCss(css, "text-decoration") || HasCss(css, "text-decoration-line");

    private static bool IsCssBold(IReadOnlyDictionary<string, string> css) =>
        TryGetCss(css, "font-weight", out string value)
        && (value.Equals("bold", StringComparison.OrdinalIgnoreCase)
            || int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int weight) && weight >= 600);

    private static bool IsCssItalic(IReadOnlyDictionary<string, string> css) =>
        TryGetCss(css, "font-style", out string value)
        && (value.Equals("italic", StringComparison.OrdinalIgnoreCase) || value.Equals("oblique", StringComparison.OrdinalIgnoreCase));

    private static bool HasDecoration(IReadOnlyDictionary<string, string> css, string decoration) =>
        (TryGetCss(css, "text-decoration-line", out string lines) || TryGetCss(css, "text-decoration", out lines))
        && lines.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries)
            .Any(value => value.Equals(decoration, StringComparison.OrdinalIgnoreCase));

    private static bool TryGetCssPoints(IReadOnlyDictionary<string, string> css, string name, out double points) {
        points = 0D;
        if (!TryGetCss(css, name, out string value)) return false;
        string normalized = value.Trim();
        double multiplier = 1D;
        if (normalized.EndsWith("px", StringComparison.OrdinalIgnoreCase)) multiplier = 0.75D;
        else if (!normalized.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) return false;
        normalized = normalized.Substring(0, normalized.Length - 2);
        return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            && parsed > 0D && (points = parsed * multiplier) > 0D;
    }

    private static string NormalizeCssFontFamily(string value) =>
        value.Split(',').FirstOrDefault()?.Trim().Trim('\'', '"') ?? string.Empty;

    private static bool TryNormalizeCssHex(string value, out string color) {
        color = value.Trim().TrimStart('#');
        if (color.Length == 3) color = string.Concat(color[0], color[0], color[1], color[1], color[2], color[2]);
        if (color.Length == 8) color = color.Substring(0, 6);
        return color.Length == 6 && color.All(Uri.IsHexDigit);
    }

    private static IEnumerable<IElement> EnumerateDirectTableRows(IElement table) {
        foreach (IElement child in table.Children) {
            if (IsElement(child, "tr")) {
                yield return child;
                continue;
            }

            if (!IsElement(child, "thead") && !IsElement(child, "tbody") && !IsElement(child, "tfoot")) {
                continue;
            }

            foreach (IElement row in child.Children.Where(candidate => IsElement(candidate, "tr"))) {
                yield return row;
            }
        }
    }

    private static bool HasDirectTableCells(IElement table) =>
        EnumerateDirectTableRows(table).Any(row => row.Children.Any(IsTableCell));

    private static bool IsTableCell(IElement element) => IsElement(element, "th") || IsElement(element, "td");

    private static int ReadSpan(
        IElement cell,
        string attributeName,
        int start,
        int maximum,
        int cellRow,
        int cellColumn,
        HtmlToExcelResult result) {
        string? rawValue = cell.GetAttribute(attributeName);
        if (string.IsNullOrWhiteSpace(rawValue)) {
            return 1;
        }

        if (!int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int span)
            || span <= 0
            || start > maximum
            || span > maximum - start + 1) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TableSpanInvalid,
                "Cell " + BuildCellReference(cellRow, cellColumn) + " contained an invalid " + attributeName + " value; a span of 1 was used.", lossKind: OfficeConversionLossKind.Approximation);
            return 1;
        }

        return span;
    }

    private static bool SpanOverlaps(HashSet<long> occupiedCells, int row, int column, int rowSpan, int columnSpan) {
        for (int currentRow = row; currentRow < row + rowSpan; currentRow++) {
            for (int currentColumn = column; currentColumn < column + columnSpan; currentColumn++) {
                if (occupiedCells.Contains(GetImportCellKey(currentRow, currentColumn))) {
                    return true;
                }
            }
        }

        return false;
    }

    private static void ReserveSpan(HashSet<long> occupiedCells, int row, int column, int rowSpan, int columnSpan) {
        for (int currentRow = row; currentRow < row + rowSpan; currentRow++) {
            for (int currentColumn = column; currentColumn < column + columnSpan; currentColumn++) {
                occupiedCells.Add(GetImportCellKey(currentRow, currentColumn));
            }
        }
    }

    private static long GetImportCellKey(int row, int column) => ((long)row << 32) | (uint)column;
}

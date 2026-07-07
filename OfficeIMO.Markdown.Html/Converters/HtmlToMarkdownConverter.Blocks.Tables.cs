using AngleSharp.Dom;
using OfficeIMO.Markdown;
using System.Globalization;

namespace OfficeIMO.Markdown.Html;

public sealed partial class HtmlToMarkdownConverter {
    private static TableBlock ConvertTableElement(IElement element, ConversionContext context) {
        var table = new TableBlock();
        bool headerWritten = false;
        var headerCells = new List<TableCell>();
        var rowCells = new List<IReadOnlyList<TableCell>>();
        var columnAlignments = new List<ColumnAlignment>();
        var columnWidthPoints = new List<double?>();
        var columnWidthWeights = new List<double?>();
        int maxExpandedColumns = context.Options.MaxTableExpandedColumns;
        CaptureColumnGroupAlignments(element, context, columnAlignments, maxExpandedColumns: maxExpandedColumns);
        CaptureColumnGroupWidths(element, context, columnWidthPoints, columnWidthWeights, maxExpandedColumns: maxExpandedColumns);
        var activeRowSpans = new List<int>();

        foreach (var row in EnumerateTableRows(element, context)) {
            var cells = row.Children
                .Where(child => HasEffectiveTagName(child, context, "TH") || HasEffectiveTagName(child, context, "TD"))
                .ToList();
            if (cells.Count == 0) {
                continue;
            }

            bool isHeaderRow = !headerWritten && cells.All(cell => HasEffectiveTagName(cell, context, "TH"));
            var renderedCells = new List<string>(cells.Count);
            var structuredCells = new List<TableCell>(cells.Count);
            int logicalColumn = 0;
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                while (logicalColumn < activeRowSpans.Count && activeRowSpans[logicalColumn] > 0) {
                    logicalColumn++;
                }

                if (logicalColumn >= maxExpandedColumns) {
                    table.SkippedColumnCount += cells.Count - cellIndex;
                    break;
                }

                var cell = cells[cellIndex];
                int parsedColumnSpan = ParseCellSpan(cell.GetAttribute("colspan"));
                int columnSpan = Math.Min(parsedColumnSpan, maxExpandedColumns - logicalColumn);
                if (columnSpan <= 0) {
                    table.SkippedColumnCount += cells.Count - cellIndex;
                    break;
                }

                if (parsedColumnSpan > columnSpan) {
                    table.SkippedColumnCount += parsedColumnSpan - columnSpan;
                }

                var cellBlocks = ConvertTableCellToBlocks(cell, context);
                ColumnAlignment cellAlignment = ParseAlignment(cell);
                var structuredCell = new TableCell(cellBlocks) {
                    Alignment = cellAlignment,
                    BackgroundColor = ParseBackgroundColor(cell),
                    TextColor = ParseTextColor(cell),
                    Bold = ParseFontWeightBold(cell),
                    Italic = ParseFontStyleItalic(cell),
                    Underline = ParseTextDecoration(cell, "underline"),
                    Strikethrough = ParseTextDecoration(cell, "line-through"),
                    ColumnSpan = columnSpan,
                    RowSpan = ParseCellSpan(cell.GetAttribute("rowspan"))
                };
                structuredCells.Add(structuredCell);
                renderedCells.Add(RenderTableCellBlocksToMarkdown(cellBlocks));
                CaptureSpannedColumnAlignment(columnAlignments, logicalColumn, structuredCell.ColumnSpan, cellAlignment, replaceExisting: isHeaderRow, maxExpandedColumns: maxExpandedColumns);
                CaptureSpannedColumnWidth(columnWidthPoints, columnWidthWeights, logicalColumn, structuredCell.ColumnSpan, ParseColumnWidth(cell), replaceExisting: false, maxExpandedColumns: maxExpandedColumns);
                UpdateActiveCellRowSpans(activeRowSpans, logicalColumn, structuredCell.ColumnSpan, structuredCell.RowSpan);
                logicalColumn += Math.Max(1, structuredCell.ColumnSpan);
            }

            DecrementActiveCellRowSpans(activeRowSpans);
            if (isHeaderRow) {
                ClampHeaderRowSpans(structuredCells);
                foreach (var value in renderedCells) {
                    table.Headers.Add(value);
                }
                headerCells.AddRange(structuredCells);
                headerWritten = true;
            } else {
                table.Rows.Add(renderedCells);
                rowCells.Add(structuredCells);
            }
        }

        if (!headerWritten && table.Rows.Count > 0) {
            var firstRow = table.Rows[0];
            table.Rows.RemoveAt(0);
            var firstStructuredRow = rowCells[0];
            rowCells.RemoveAt(0);
            foreach (var value in firstRow) {
                table.Headers.Add(value);
            }
            ClampHeaderRowSpans(firstStructuredRow);
            headerCells.AddRange(firstStructuredRow);
        }

        ApplyColumnAlignments(table, columnAlignments);
        ApplyColumnWidths(table, columnWidthPoints, columnWidthWeights);
        table.SetStructuredCells(headerCells, rowCells, table.ComputeContentSignature());

        return table;
    }

    private static void ClampHeaderRowSpans(IReadOnlyList<TableCell> cells) {
        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
            if (cells[cellIndex].RowSpan > 1) {
                cells[cellIndex].RowSpan = 1;
            }
        }
    }

    private static void CaptureColumnAlignment(List<ColumnAlignment> alignments, int columnIndex, ColumnAlignment alignment, bool replaceExisting) {
        while (alignments.Count <= columnIndex) {
            alignments.Add(ColumnAlignment.None);
        }

        if (alignment == ColumnAlignment.None) {
            return;
        }

        if (replaceExisting || alignments[columnIndex] == ColumnAlignment.None) {
            alignments[columnIndex] = alignment;
        }
    }

    private static void CaptureColumnGroupAlignments(IElement table, ConversionContext context, List<ColumnAlignment> alignments, int maxExpandedColumns) {
        int columnIndex = 0;
        foreach (var child in table.Children) {
            if (HasEffectiveTagName(child, context, "COLGROUP")) {
                var colElements = child.Children
                    .Where(col => HasEffectiveTagName(col, context, "COL"))
                    .ToList();
                var groupAlignment = ParseAlignment(child);
                if (colElements.Count == 0) {
                    columnIndex = CaptureSpannedColumnAlignment(alignments, columnIndex, child, groupAlignment, replaceExisting: false, maxExpandedColumns: maxExpandedColumns);
                    continue;
                }

                foreach (var col in colElements) {
                    var columnAlignment = ParseAlignment(col);
                    if (columnAlignment == ColumnAlignment.None) {
                        columnAlignment = groupAlignment;
                    }

                    columnIndex = CaptureSpannedColumnAlignment(alignments, columnIndex, col, columnAlignment, replaceExisting: false, maxExpandedColumns: maxExpandedColumns);
                }

                continue;
            }

            if (HasEffectiveTagName(child, context, "COL")) {
                columnIndex = CaptureSpannedColumnAlignment(alignments, columnIndex, child, ParseAlignment(child), replaceExisting: false, maxExpandedColumns: maxExpandedColumns);
            }
        }
    }

    private static int CaptureSpannedColumnAlignment(List<ColumnAlignment> alignments, int columnIndex, IElement columnElement, ColumnAlignment alignment, bool replaceExisting, int maxExpandedColumns) {
        int span = ParseColumnSpan(columnElement.GetAttribute("span"));
        return CaptureSpannedColumnAlignment(alignments, columnIndex, span, alignment, replaceExisting, maxExpandedColumns: maxExpandedColumns);
    }

    private static int CaptureSpannedColumnAlignment(List<ColumnAlignment> alignments, int columnIndex, int span, ColumnAlignment alignment, bool replaceExisting, int maxExpandedColumns) {
        int boundedSpan = Math.Min(span, Math.Max(0, maxExpandedColumns - columnIndex));
        for (int offset = 0; offset < boundedSpan; offset++) {
            CaptureColumnAlignment(alignments, columnIndex + offset, alignment, replaceExisting);
        }

        return columnIndex + boundedSpan;
    }

    private static void CaptureColumnGroupWidths(IElement table, ConversionContext context, List<double?> widthPoints, List<double?> widthWeights, int maxExpandedColumns) {
        int columnIndex = 0;
        foreach (var child in table.Children) {
            if (HasEffectiveTagName(child, context, "COLGROUP")) {
                var colElements = child.Children
                    .Where(col => HasEffectiveTagName(col, context, "COL"))
                    .ToList();
                ColumnWidthHint groupWidth = ParseColumnWidth(child);
                if (colElements.Count == 0) {
                    columnIndex = CaptureSpannedColumnWidth(widthPoints, widthWeights, columnIndex, child, groupWidth, replaceExisting: false, maxExpandedColumns: maxExpandedColumns);
                    continue;
                }

                foreach (var col in colElements) {
                    ColumnWidthHint columnWidth = ParseColumnWidth(col);
                    if (!columnWidth.HasValue) {
                        columnWidth = groupWidth;
                    }

                    columnIndex = CaptureSpannedColumnWidth(widthPoints, widthWeights, columnIndex, col, columnWidth, replaceExisting: false, maxExpandedColumns: maxExpandedColumns);
                }

                continue;
            }

            if (HasEffectiveTagName(child, context, "COL")) {
                columnIndex = CaptureSpannedColumnWidth(widthPoints, widthWeights, columnIndex, child, ParseColumnWidth(child), replaceExisting: false, maxExpandedColumns: maxExpandedColumns);
            }
        }
    }

    private static int CaptureSpannedColumnWidth(List<double?> widthPoints, List<double?> widthWeights, int columnIndex, IElement columnElement, ColumnWidthHint width, bool replaceExisting, int maxExpandedColumns) {
        int span = ParseColumnSpan(columnElement.GetAttribute("span"));
        return CaptureSpannedColumnWidth(widthPoints, widthWeights, columnIndex, span, width, replaceExisting, maxExpandedColumns: maxExpandedColumns);
    }

    private static int CaptureSpannedColumnWidth(List<double?> widthPoints, List<double?> widthWeights, int columnIndex, int span, ColumnWidthHint width, bool replaceExisting, int maxExpandedColumns) {
        int boundedSpan = Math.Min(span, Math.Max(0, maxExpandedColumns - columnIndex));
        for (int offset = 0; offset < boundedSpan; offset++) {
            CaptureColumnWidth(widthPoints, widthWeights, columnIndex + offset, width, replaceExisting);
        }

        return columnIndex + boundedSpan;
    }

    private static void CaptureColumnWidth(List<double?> widthPoints, List<double?> widthWeights, int columnIndex, ColumnWidthHint width, bool replaceExisting) {
        if (!width.HasValue) {
            return;
        }

        while (widthPoints.Count <= columnIndex) {
            widthPoints.Add(null);
        }

        while (widthWeights.Count <= columnIndex) {
            widthWeights.Add(null);
        }

        if (!replaceExisting && (widthPoints[columnIndex].HasValue || widthWeights[columnIndex].HasValue)) {
            return;
        }

        widthPoints[columnIndex] = width.Points;
        widthWeights[columnIndex] = width.Weight;
    }

    private static int ParseColumnSpan(string? rawSpan) {
        if (!int.TryParse(rawSpan, out int span) || span < 1) {
            return 1;
        }

        return Math.Min(span, 512);
    }

    private static int ParseCellSpan(string? rawSpan) {
        if (!int.TryParse(rawSpan, out int span) || span < 1) {
            return 1;
        }

        return Math.Min(span, 512);
    }

    private static void UpdateActiveCellRowSpans(List<int> activeRowSpans, int logicalColumn, int columnSpan, int rowSpan) {
        if (rowSpan <= 1) {
            return;
        }

        int span = Math.Max(1, columnSpan);
        while (activeRowSpans.Count < logicalColumn + span) {
            activeRowSpans.Add(0);
        }

        for (int offset = 0; offset < span; offset++) {
            int column = logicalColumn + offset;
            activeRowSpans[column] = Math.Max(activeRowSpans[column], rowSpan);
        }
    }

    private static void DecrementActiveCellRowSpans(List<int> activeRowSpans) {
        for (int column = 0; column < activeRowSpans.Count; column++) {
            if (activeRowSpans[column] > 0) {
                activeRowSpans[column]--;
            }
        }
    }

    private static void ApplyColumnWidths(TableBlock table, List<double?> widthPoints, List<double?> widthWeights) {
        int columnCount = table.Headers.Count;
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            columnCount = Math.Max(columnCount, table.Rows[rowIndex].Count);
        }

        bool hasFixedWidths = false;
        table.ColumnWidthPoints.Clear();
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            double? width = columnIndex < widthPoints.Count ? widthPoints[columnIndex] : null;
            table.ColumnWidthPoints.Add(width);
            hasFixedWidths |= width.HasValue;
        }

        if (!hasFixedWidths) {
            table.ColumnWidthPoints.Clear();
        }

        bool hasWeights = false;
        table.ColumnWidthWeights.Clear();
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            double weight = columnIndex < widthWeights.Count && widthWeights[columnIndex].HasValue
                ? widthWeights[columnIndex]!.Value
                : 1D;
            table.ColumnWidthWeights.Add(weight);
            if (Math.Abs(weight - 1D) > 0.001) {
                hasWeights = true;
            }
        }

        if (!hasWeights) {
            table.ColumnWidthWeights.Clear();
        }
    }

    private static void ApplyColumnAlignments(TableBlock table, List<ColumnAlignment> alignments) {
        int columnCount = table.Headers.Count;
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            columnCount = Math.Max(columnCount, table.Rows[rowIndex].Count);
        }

        table.Alignments.Clear();
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            table.Alignments.Add(columnIndex < alignments.Count ? alignments[columnIndex] : ColumnAlignment.None);
        }
    }

    private static IEnumerable<IElement> EnumerateTableRows(IElement table, ConversionContext context) {
        foreach (var child in table.Children) {
            if (HasEffectiveTagName(child, context, "TR")) {
                yield return child;
                continue;
            }

            if (!HasEffectiveTagName(child, context, "THEAD")
                && !HasEffectiveTagName(child, context, "TBODY")
                && !HasEffectiveTagName(child, context, "TFOOT")) {
                continue;
            }

            foreach (var row in child.Children.Where(row => HasEffectiveTagName(row, context, "TR"))) {
                yield return row;
            }
        }
    }

    private static ColumnAlignment ParseAlignment(IElement cell) {
        var alignment = ParseAlignmentValue(cell.GetAttribute("align"));
        if (alignment != ColumnAlignment.None) {
            return alignment;
        }

        alignment = ParseAlignmentValue(TryGetStyleDeclarationValue(cell.GetAttribute("style"), "text-align"));
        if (alignment != ColumnAlignment.None) {
            return alignment;
        }

        return ParseAlignmentClassTokens(cell.GetAttribute("class"));
    }

    private static ColumnWidthHint ParseColumnWidth(IElement element) {
        ColumnWidthHint width = ParseColumnWidthValue(element.GetAttribute("width"));
        if (width.HasValue) {
            return width;
        }

        return ParseColumnWidthValue(TryGetStyleDeclarationValue(element.GetAttribute("style"), "width"));
    }

    private static string? ParseBackgroundColor(IElement element) {
        string? color = NormalizeCssColor(TryGetStyleDeclarationValue(element.GetAttribute("style"), "background-color"));
        if (color != null) {
            return color;
        }

        return NormalizeCssBackgroundColor(TryGetStyleDeclarationValue(element.GetAttribute("style"), "background"));
    }

    private static string? ParseTextColor(IElement element) {
        return NormalizeCssColor(TryGetStyleDeclarationValue(element.GetAttribute("style"), "color"));
    }

    private static bool ParseFontWeightBold(IElement element) {
        string? value = TryGetStyleDeclarationValue(element.GetAttribute("style"), "font-weight");
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string fontWeight = value!.Trim();
        int importantIndex = fontWeight.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
        if (importantIndex >= 0) {
            fontWeight = fontWeight.Substring(0, importantIndex).Trim();
        }

        if (fontWeight.Equals("bold", StringComparison.OrdinalIgnoreCase) || fontWeight.Equals("bolder", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        return int.TryParse(fontWeight, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numericWeight) && numericWeight >= 600;
    }

    private static bool ParseFontStyleItalic(IElement element) {
        string? value = TryGetStyleDeclarationValue(element.GetAttribute("style"), "font-style");
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string fontStyle = value!.Trim();
        int importantIndex = fontStyle.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
        if (importantIndex >= 0) {
            fontStyle = fontStyle.Substring(0, importantIndex).Trim();
        }

        return fontStyle.Equals("italic", StringComparison.OrdinalIgnoreCase) || fontStyle.Equals("oblique", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ParseTextDecoration(IElement element, string decoration) {
        string? value = TryGetStyleDeclarationValue(element.GetAttribute("style"), "text-decoration");
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string textDecoration = value!.Trim();
        int importantIndex = textDecoration.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
        if (importantIndex >= 0) {
            textDecoration = textDecoration.Substring(0, importantIndex).Trim();
        }

        return textDecoration.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Any(token => token.Equals(decoration, StringComparison.OrdinalIgnoreCase));
    }

    private static string? NormalizeCssBackgroundColor(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        string candidate = value!.Trim();
        if (candidate.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase) || candidate.StartsWith("rgba(", StringComparison.OrdinalIgnoreCase)) {
            int end = candidate.IndexOf(')');
            return end > 0 ? NormalizeCssColor(candidate.Substring(0, end + 1)) : null;
        }

        int separator = candidate.IndexOfAny(new[] { ' ', '\t', '\r', '\n' });
        if (separator > 0) {
            candidate = candidate.Substring(0, separator);
        }

        return NormalizeCssColor(candidate);
    }

    private static string? NormalizeCssColor(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        string color = value!.Trim();
        int importantIndex = color.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
        if (importantIndex >= 0) {
            color = color.Substring(0, importantIndex).Trim();
        }

        if (color.Equals("transparent", StringComparison.OrdinalIgnoreCase)) {
            return null;
        }

        if (TryNormalizeRgbColor(color, out string? normalizedRgb)) {
            return normalizedRgb;
        }

        if (color.StartsWith("#", StringComparison.Ordinal) && (color.Length == 4 || color.Length == 5 || color.Length == 7 || color.Length == 9)) {
            return color.ToLowerInvariant();
        }

        if (color.All(static c => char.IsLetter(c))) {
            return color;
        }

        return null;
    }

    private static bool TryNormalizeRgbColor(string color, out string? normalized) {
        normalized = null;
        bool rgba = color.StartsWith("rgba(", StringComparison.OrdinalIgnoreCase);
        if (!rgba && !color.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        int start = color.IndexOf('(');
        int end = color.LastIndexOf(')');
        if (start < 0 || end <= start) {
            return false;
        }

        string[] parts = color.Substring(start + 1, end - start - 1).Split(',');
        if ((!rgba && parts.Length != 3) || (rgba && parts.Length != 4)) {
            return false;
        }

        if (!TryParseCssByte(parts[0], out byte r) ||
            !TryParseCssByte(parts[1], out byte g) ||
            !TryParseCssByte(parts[2], out byte b)) {
            return false;
        }

        if (rgba && TryParseCssAlpha(parts[3], out double alpha) && alpha <= 0D) {
            return true;
        }

        normalized = "#" + r.ToString("x2", CultureInfo.InvariantCulture) + g.ToString("x2", CultureInfo.InvariantCulture) + b.ToString("x2", CultureInfo.InvariantCulture);
        return true;
    }

    private static bool TryParseCssByte(string value, out byte result) {
        result = 0;
        string trimmed = value.Trim();
        if (trimmed.EndsWith("%", StringComparison.Ordinal)) {
            string percentText = trimmed.Substring(0, trimmed.Length - 1).Trim();
            if (!double.TryParse(percentText, NumberStyles.Float, CultureInfo.InvariantCulture, out double percent)) {
                return false;
            }

            result = ClampCssByte(percent * 255D / 100D);
            return true;
        }

        if (!double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
            return false;
        }

        result = ClampCssByte(number);
        return true;
    }

    private static bool TryParseCssAlpha(string value, out double result) {
        string trimmed = value.Trim();
        if (trimmed.EndsWith("%", StringComparison.Ordinal)) {
            if (!double.TryParse(trimmed.Substring(0, trimmed.Length - 1).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out result)) {
                return false;
            }

            result /= 100D;
            return true;
        }

        return double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out result);
    }

    private static byte ClampCssByte(double value) {
        if (value <= 0D || double.IsNaN(value)) {
            return 0;
        }

        if (value >= 255D || double.IsInfinity(value)) {
            return 255;
        }

        return (byte)Math.Round(value);
    }

    private static ColumnWidthHint ParseColumnWidthValue(string? rawWidth) {
        if (string.IsNullOrWhiteSpace(rawWidth)) {
            return ColumnWidthHint.None;
        }

        string value = rawWidth!.Trim();
        int importantIndex = value.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
        if (importantIndex >= 0) {
            value = value.Substring(0, importantIndex).Trim();
        }

        if (value.EndsWith("%", StringComparison.Ordinal)) {
            string numeric = value.Substring(0, value.Length - 1).Trim();
            if (TryParsePositiveFiniteDouble(numeric, out double percentage)) {
                return ColumnWidthHint.FromWeight(percentage);
            }
        }

        string unit = string.Empty;
        int unitStart = value.Length;
        while (unitStart > 0 && (char.IsLetter(value[unitStart - 1]) || value[unitStart - 1] == '%')) {
            unitStart--;
        }

        if (unitStart < value.Length) {
            unit = value.Substring(unitStart).Trim().ToLowerInvariant();
            value = value.Substring(0, unitStart).Trim();
        }

        if (!TryParsePositiveFiniteDouble(value, out double numericValue)) {
            return ColumnWidthHint.None;
        }

        switch (unit) {
            case "":
            case "px":
                return ColumnWidthHint.FromPoints(numericValue * 0.75D);
            case "pt":
                return ColumnWidthHint.FromPoints(numericValue);
            case "in":
                return ColumnWidthHint.FromPoints(numericValue * 72D);
            case "cm":
                return ColumnWidthHint.FromPoints(numericValue * 72D / 2.54D);
            case "mm":
                return ColumnWidthHint.FromPoints(numericValue * 72D / 25.4D);
            case "pc":
                return ColumnWidthHint.FromPoints(numericValue * 12D);
            default:
                return ColumnWidthHint.None;
        }
    }

    private static bool TryParsePositiveFiniteDouble(string value, out double result) {
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out result)
            && result > 0D
            && !double.IsNaN(result)
            && !double.IsInfinity(result)) {
            return true;
        }

        result = 0D;
        return false;
    }

    private static ColumnAlignment ParseAlignmentValue(string? rawAlignment) {
        if (string.IsNullOrWhiteSpace(rawAlignment)) {
            return ColumnAlignment.None;
        }

        switch (rawAlignment!.Trim().ToLowerInvariant()) {
            case "left":
                return ColumnAlignment.Left;
            case "center":
                return ColumnAlignment.Center;
            case "right":
                return ColumnAlignment.Right;
            default:
                return ColumnAlignment.None;
        }
    }

    private static ColumnAlignment ParseAlignmentClassTokens(string? classValue) {
        if (string.IsNullOrWhiteSpace(classValue)) {
            return ColumnAlignment.None;
        }

        foreach (string rawToken in classValue!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)) {
            string token = rawToken.Trim().Replace('_', '-').ToLowerInvariant();
            switch (token) {
                case "left":
                case "align-left":
                case "alignleft":
                case "left-align":
                case "text-left":
                case "ta-left":
                case "has-text-left":
                case "is-left":
                    return ColumnAlignment.Left;
                case "center":
                case "align-center":
                case "aligncenter":
                case "center-align":
                case "text-center":
                case "ta-center":
                case "has-text-center":
                case "has-text-centered":
                case "is-center":
                case "is-centered":
                    return ColumnAlignment.Center;
                case "right":
                case "align-right":
                case "alignright":
                case "right-align":
                case "text-right":
                case "ta-right":
                case "has-text-right":
                case "is-right":
                    return ColumnAlignment.Right;
            }
        }

        return ColumnAlignment.None;
    }

    private static IReadOnlyList<IMarkdownBlock> ConvertTableCellToBlocks(IElement cell, ConversionContext context) {
        if (HasDirectBlockChildren(cell, context)) {
            return ConvertNodesToBlocks(cell.ChildNodes, context);
        }

        var inlineSequence = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(cell.ChildNodes, context));
        if (!HasVisibleInlineContent(inlineSequence)) {
            return Array.Empty<IMarkdownBlock>();
        }

        return new IMarkdownBlock[] { new ParagraphBlock(inlineSequence) };
    }

    private static string RenderTableCellBlocksToMarkdown(IReadOnlyList<IMarkdownBlock> blocks) {
        if (blocks == null || blocks.Count == 0) {
            return string.Empty;
        }

        return new TableCell(blocks).Markdown.Replace("  \n", "<br>");
    }

    private sealed class ColumnWidthHint {
        private ColumnWidthHint(double? points, double? weight) {
            Points = points;
            Weight = weight;
        }

        public static ColumnWidthHint None { get; } = new ColumnWidthHint(null, null);

        public double? Points { get; }

        public double? Weight { get; }

        public bool HasValue => Points.HasValue || Weight.HasValue;

        public static ColumnWidthHint FromPoints(double points) => new ColumnWidthHint(points, null);

        public static ColumnWidthHint FromWeight(double weight) => new ColumnWidthHint(null, weight);
    }

}

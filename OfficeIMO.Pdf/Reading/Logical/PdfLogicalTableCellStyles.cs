using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Visible styling of one reconstructed table cell, taken from the source text runs and cell fill.</summary>
internal sealed class PdfLogicalTableCellStyle {
    internal PdfLogicalTableCellStyle(OfficeColor? fillColor, OfficeColor? textColor, bool isBold, bool isItalic, double? fontSize, string? fontName) {
        FillColor = fillColor;
        TextColor = textColor;
        IsBold = isBold;
        IsItalic = isItalic;
        FontSize = fontSize;
        FontName = fontName;
    }

    /// <summary>Opaque background painted under the cell, or null when the cell is unfilled or white.</summary>
    internal OfficeColor? FillColor { get; }

    /// <summary>Dominant text colour of the cell runs.</summary>
    internal OfficeColor? TextColor { get; }

    internal bool IsBold { get; }

    internal bool IsItalic { get; }

    internal double? FontSize { get; }

    /// <summary>PostScript base font name of the dominant run.</summary>
    internal string? FontName { get; }

    internal bool HasVisibleStyle => FillColor.HasValue || IsBold;
}

/// <summary>Reads cell fills and text styling for logical tables from the page that produced them.</summary>
internal static class PdfLogicalTableCellStyleAnalysis {
    private readonly struct FillRectangle {
        internal FillRectangle(double left, double top, double right, double bottom, OfficeColor color, double paintOrder) {
            Left = left; Top = top; Right = right; Bottom = bottom; Color = color; PaintOrder = paintOrder;
        }
        internal double Left { get; }
        internal double Top { get; }
        internal double Right { get; }
        internal double Bottom { get; }
        internal OfficeColor Color { get; }
        internal double PaintOrder { get; }
        internal bool Contains(double x, double y) => x >= Left && x <= Right && y >= Top && y <= Bottom;
    }

    private sealed class CellAccumulator {
        internal readonly Dictionary<OfficeColor, double> Colors = new Dictionary<OfficeColor, double>();
        internal readonly Dictionary<string, double> Fonts = new Dictionary<string, double>(StringComparer.Ordinal);
        internal double BoldWeight;
        internal double ItalicWeight;
        internal double TotalWeight;
        internal readonly List<double> FontSizes = new List<double>();
        internal double Bottom = double.MaxValue;
        internal double Top = double.MinValue;
        internal bool HasText => TotalWeight > 0D;
    }

    internal static void Assign(PdfReadPage page, IReadOnlyList<PdfLogicalTable> tables, CancellationToken cancellationToken) {
        if (tables.Count == 0 || !tables.Any(static table => table.SourceLines.Count > 0 && table.Columns.Count > 0 && table.Rows.Count > 0)) return;
        List<FillRectangle> fills;
        try {
            fills = CollectFills(page, cancellationToken);
        } catch (Exception exception) when (exception is not OperationCanceledException && exception is not PdfReadLimitException) {
            // Styling is best effort; the table content itself is already reconstructed.
            fills = new List<FillRectangle>();
        }
        for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfLogicalTable table = tables[tableIndex];
            if (table.SourceLines.Count == 0 || table.Columns.Count == 0 || table.Rows.Count == 0) continue;
            table.CellStyles = Analyze(page, table, fills, cancellationToken);
        }
    }

    private static List<FillRectangle> CollectFills(PdfReadPage page, CancellationToken cancellationToken) {
        (double width, double height) = page.GetVisualPageSize();
        double pageArea = Math.Max(1D, width * height);
        var result = new List<FillRectangle>();
        foreach (PdfPageVisualPrimitive primitive in page.GetIdentityVisualPrimitives(cancellationToken)) {
            if (primitive.Kind is not (PdfPageVisualPrimitiveKind.Rectangle or PdfPageVisualPrimitiveKind.Path)) continue;
            if (!primitive.FillColor.HasValue || primitive.FillGradient != null || primitive.FillRadialGradient != null) continue;
            if (primitive.FillOpacity is double opacity && opacity < 0.5D) continue;
            OfficeColor color = primitive.FillColor.Value;
            if (color.A < 128) continue;
            double w = Math.Abs(primitive.Width), h = Math.Abs(primitive.Height);
            if (w < 1D || h < 1D || w * h >= pageArea * 0.8D) continue;
            double left = Math.Min(primitive.X, primitive.X + primitive.Width);
            double top = Math.Min(primitive.Y, primitive.Y + primitive.Height);
            result.Add(new FillRectangle(left, top, left + w, top + h, color, primitive.PaintOrder));
        }
        return result;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<IReadOnlyList<PdfLogicalTableCellStyle?>> Analyze(
        PdfReadPage page,
        PdfLogicalTable table,
        List<FillRectangle> fills,
        CancellationToken cancellationToken) {
        int rowCount = table.Rows.Count;
        int columnCount = table.Columns.Count;
        var cells = new CellAccumulator?[rowCount, columnCount];
        PdfUnderstandingLine[] lines = table.SourceLines.OrderByDescending(static line => line.BaselineY).ThenBy(static line => line.XStart).ToArray();
        int rowCursor = 0;
        foreach (PdfUnderstandingLine line in lines) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (PdfUnderstandingWord word in line.Words) {
                string text = word.Text.Trim();
                if (text.Length == 0) continue;
                int column = FindColumn(table, (word.XStart + word.XEnd) / 2D);
                int row = FindRow(table, column, text, rowCursor);
                if (row < 0) continue;
                rowCursor = row;
                CellAccumulator cell = cells[row, column] ??= new CellAccumulator();
                double fontSize = Math.Max(1D, word.FontSize);
                cell.Bottom = Math.Min(cell.Bottom, word.BaselineY - fontSize * 0.25D);
                cell.Top = Math.Max(cell.Top, word.BaselineY + fontSize * 0.8D);
                IReadOnlyList<PdfTextSpan> runs = word.SourceRuns;
                double weight = Math.Max(1, text.Length);
                PdfTextSpan? run = runs.Count > 0 ? runs[0] : null;
                cell.TotalWeight += weight;
                cell.FontSizes.Add(fontSize);
                if (run == null) continue;
                if (run.IsBold) cell.BoldWeight += weight;
                if (run.IsItalic) cell.ItalicWeight += weight;
                if (run.Color.HasValue) cell.Colors[run.Color.Value] = cell.Colors.TryGetValue(run.Color.Value, out double existing) ? existing + weight : weight;
                if (!string.IsNullOrWhiteSpace(run.BaseFont)) cell.Fonts[run.BaseFont!] = cell.Fonts.TryGetValue(run.BaseFont!, out double font) ? font + weight : weight;
            }
        }

        var rows = new IReadOnlyList<PdfLogicalTableCellStyle?>[rowCount];
        for (int row = 0; row < rowCount; row++) {
            double rowBottom = double.MaxValue, rowTop = double.MinValue;
            for (int column = 0; column < columnCount; column++) {
                CellAccumulator? cell = cells[row, column];
                if (cell?.HasText != true) continue;
                rowBottom = Math.Min(rowBottom, cell!.Bottom);
                rowTop = Math.Max(rowTop, cell.Top);
            }
            var styles = new PdfLogicalTableCellStyle?[columnCount];
            for (int column = 0; column < columnCount; column++) {
                CellAccumulator? cell = cells[row, column];
                PdfLogicalTableColumn source = table.Columns[column];
                double left = Math.Min(source.From, source.To), right = Math.Max(source.From, source.To);
                double bottom = cell?.HasText == true ? cell.Bottom : rowBottom;
                double top = cell?.HasText == true ? cell.Top : rowTop;
                OfficeColor? fill = null;
                if (top > bottom && right > left) {
                    PdfVisualBounds visual = page.TransformBoundsToVisual(left, bottom, right, top);
                    fill = FindFill(fills, (visual.Left + visual.Right) / 2D, (visual.Top + visual.Bottom) / 2D);
                }
                if (cell?.HasText != true) {
                    styles[column] = fill.HasValue ? new PdfLogicalTableCellStyle(fill, null, false, false, null, null) : null;
                    continue;
                }
                double[] sizes = cell!.FontSizes.OrderBy(static size => size).ToArray();
                OfficeColor? textColor = cell.Colors.Count == 0 ? null : cell.Colors.OrderByDescending(static pair => pair.Value).First().Key;
                string? fontName = cell.Fonts.Count == 0 ? null : cell.Fonts.OrderByDescending(static pair => pair.Value).First().Key;
                styles[column] = new PdfLogicalTableCellStyle(
                    fill,
                    textColor,
                    cell.BoldWeight * 2D > cell.TotalWeight,
                    cell.ItalicWeight * 2D > cell.TotalWeight,
                    sizes[sizes.Length / 2],
                    fontName);
            }
            rows[row] = Array.AsReadOnly(styles);
        }
        return Array.AsReadOnly(rows);
    }

    private static OfficeColor? FindFill(List<FillRectangle> fills, double x, double y) {
        FillRectangle? best = null;
        for (int index = 0; index < fills.Count; index++) {
            FillRectangle fill = fills[index];
            if (!fill.Contains(x, y)) continue;
            if (best == null || fill.PaintOrder >= best.Value.PaintOrder) best = fill;
        }
        if (best == null) return null;
        OfficeColor color = best.Value.Color;
        // White fills are the page colour; they carry no visible styling of their own.
        return color.R >= 250 && color.G >= 250 && color.B >= 250 ? null : color;
    }

    private static int FindColumn(PdfLogicalTable table, double x) {
        int best = 0;
        double bestDistance = double.MaxValue;
        for (int index = 0; index < table.Columns.Count; index++) {
            PdfLogicalTableColumn column = table.Columns[index];
            double left = Math.Min(column.From, column.To), right = Math.Max(column.From, column.To);
            double distance = x < left ? left - x : x > right ? x - right : 0D;
            if (distance < bestDistance) {
                best = index;
                bestDistance = distance;
            }
        }
        return best;
    }

    private static int FindRow(PdfLogicalTable table, int column, string text, int start) {
        // Rows are reconstructed top to bottom, so search forward from the current row first.
        for (int row = start; row < table.Rows.Count; row++) {
            if (CellContains(table, row, column, text)) return row;
        }
        for (int row = start - 1; row >= 0; row--) {
            if (CellContains(table, row, column, text)) return row;
        }
        return -1;
    }

    private static bool CellContains(PdfLogicalTable table, int row, int column, string text) {
        IReadOnlyList<string> cells = table.Rows[row];
        return column < cells.Count && cells[column] != null && ContainsOrdinal(cells[column], text);
    }

    private static bool ContainsOrdinal(string value, string candidate) {
#if NETSTANDARD2_0 || NETFRAMEWORK
        return value.IndexOf(candidate, StringComparison.Ordinal) >= 0;
#else
        return value.Contains(candidate, StringComparison.Ordinal);
#endif
    }
}

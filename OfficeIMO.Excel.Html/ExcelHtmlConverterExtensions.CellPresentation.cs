using OfficeIMO.Drawing;

namespace OfficeIMO.Excel.Html;

public static partial class ExcelHtmlConverterExtensions {
    /// <summary>
    /// Resolved worksheet presentation for one exported table window: the same display text,
    /// table-style header fills and conditional formats the visual exporters use.
    /// </summary>
    private sealed class SemanticCellPresentation {
        private readonly Dictionary<long, ExcelVisualCell> _cells;
        private readonly Dictionary<long, ExcelVisualConditionalDataBar> _dataBars;

        private SemanticCellPresentation(Dictionary<long, ExcelVisualCell> cells, Dictionary<long, ExcelVisualConditionalDataBar> dataBars) {
            _cells = cells;
            _dataBars = dataBars;
        }

        internal bool TryGetCell(int row, int column, out ExcelVisualCell cell) => _cells.TryGetValue(Key(row, column), out cell!);

        internal bool TryGetDataBar(int row, int column, out ExcelVisualConditionalDataBar bar) => _dataBars.TryGetValue(Key(row, column), out bar!);

        private static long Key(int row, int column) => ((long)row << 20) | (uint)column;

        internal static SemanticCellPresentation? TryCreate(ExcelSheet sheet, int firstRow, int firstColumn, int rowCount, int columnCount) {
            if (rowCount <= 0 || columnCount <= 0) return null;
            var options = new ExcelImageExportOptions {
                IncludeHidden = true,
                IncludeImages = false,
                IncludeCharts = false,
                IncludeDrawingObjects = false,
                ShowHyperlinkHints = false
            };
            if ((long)rowCount * columnCount > options.MaximumRenderedCells) return null;
            string range = ExcelA1.CellReference(firstRow, firstColumn) + ":" + ExcelA1.CellReference(firstRow + rowCount - 1, firstColumn + columnCount - 1);
            ExcelRangeVisualInfo snapshot;
            try {
                snapshot = ExcelRangeVisualSnapshotBuilder.Build(sheet, range, options);
            } catch (InvalidOperationException) {
                return null;
            } catch (ArgumentException) {
                return null;
            }

            var cells = new Dictionary<long, ExcelVisualCell>(snapshot.Cells.Count);
            foreach (ExcelVisualCell cell in snapshot.Cells) cells[Key(cell.Row, cell.Column)] = cell;
            var bars = new Dictionary<long, ExcelVisualConditionalDataBar>();
            foreach (ExcelVisualConditionalDataBar bar in snapshot.ConditionalDataBars) bars[Key(bar.Row, bar.Column)] = bar;
            return new SemanticCellPresentation(cells, bars);
        }
    }

    /// <summary>Box-level CSS for a semantic cell: fill, data bar and spreadsheet alignment.</summary>
    private static string? CreateCellBoxCss(ExcelCellStyleInfo style, ExcelVisualCellValueKind? valueKind, ExcelVisualConditionalDataBar? dataBar) {
        var css = new StringBuilder();
        string? fill = ResolveCellFillHex(style);
        if (fill != null) AppendCss(css, "background-color", "#" + fill);
        if (dataBar != null && ToCssHex(dataBar.ColorArgb) is string barColor) {
            double start = Clamp01(dataBar.StartRatio) * 100D;
            double end = Clamp01(dataBar.StartRatio + dataBar.Ratio) * 100D;
            string Percent(double value) => value.ToString("0.##", CultureInfo.InvariantCulture) + "%";
            AppendCss(css, "background-image", "linear-gradient(90deg,transparent " + Percent(start) + ",#" + barColor + " " + Percent(start) + ",#" + barColor + " " + Percent(end) + ",transparent " + Percent(end) + ")");
        }

        AppendCss(css, "text-align", ResolveTextAlign(style.HorizontalAlignment, valueKind));
        AppendCss(css, "vertical-align", style.VerticalAlignment switch {
            "top" => "top",
            "center" => "middle",
            "bottom" => "bottom",
            _ => null
        });
        if (style.WrapText) AppendCss(css, "white-space", "pre-wrap");
        if (style.TextIndent is uint indent && indent > 0U) {
            AppendCss(css, "padding-left", (indent * 9U).ToString(CultureInfo.InvariantCulture) + "pt");
        }

        return css.Length == 0 ? null : css.ToString();
    }

    private static string? ResolveTextAlign(string? horizontal, ExcelVisualCellValueKind? valueKind) {
        switch (horizontal) {
            case "left":
            case "fill":
                return "left";
            case "center":
            case "centerContinuous":
                return "center";
            case "right":
                return "right";
            case "justify":
            case "distributed":
                return "justify";
        }

        // Excel's General alignment: numbers and dates to the right, logical and error values centred.
        return valueKind switch {
            ExcelVisualCellValueKind.Number or ExcelVisualCellValueKind.Date => "right",
            ExcelVisualCellValueKind.Boolean or ExcelVisualCellValueKind.Error => "center",
            _ => null
        };
    }

    private static string? ResolveCellFillHex(ExcelCellStyleInfo style) {
        if (style.FillPatternType == null) return null;
        if (style.FillPatternType == "gradient") return ToCssHex(style.FillGradientStartColorArgb);
        if (style.FillPatternType == "solid") return style.FillColorHex ?? ToCssHex(style.FillPatternForegroundColorArgb);
        // Pattern fills are approximated by their dominant foreground colour.
        return ToCssHex(style.FillPatternForegroundColorArgb) ?? style.FillColorHex;
    }

    /// <summary>
    /// White or near-white text without a fill disappears on the page background. That happens
    /// when the fill comes from a feature HTML cannot reproduce, so keep the text readable.
    /// </summary>
    private static bool IsInvisibleOnPage(ExcelCellStyleInfo style, bool hasBackground) {
        if (hasBackground || style.FontColorHex is not string color || color.Length != 6) return false;
        if (!int.TryParse(color, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int rgb)) return false;
        int r = (rgb >> 16) & 0xFF, g = (rgb >> 8) & 0xFF, b = rgb & 0xFF;
        return r >= 0xF0 && g >= 0xF0 && b >= 0xF0;
    }

    private static string? ToCssHex(string? argb) {
        if (string.IsNullOrWhiteSpace(argb)) return null;
        string value = argb!.Trim().TrimStart('#');
        if (value.Length == 8) value = value.Substring(2);
        return value.Length == 6 ? value.ToUpperInvariant() : null;
    }

    private static double Clamp01(double value) => double.IsNaN(value) ? 0D : Math.Max(0D, Math.Min(1D, value));
}

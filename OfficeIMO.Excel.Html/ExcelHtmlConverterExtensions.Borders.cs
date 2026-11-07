using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class ExcelHtmlConverterExtensions {
    /// <summary>Projects an already-resolved native edge into the closest CSS border representation.</summary>
    private static void AppendCellBorderCss(StringBuilder css, string edge, ExcelBorderSideInfo? border) {
        if (border == null) return;
        string? style = border.Style?.Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(style) || style == "none") {
            AppendCss(css, "border-" + edge, "none");
            return;
        }

        string width = style switch {
            "hair" => "0.5px",
            "medium" or "mediumdashed" or "mediumdashdot" or "mediumdashdotdot" or "slantdashdot" => "2px",
            "thick" or "double" => "3px",
            _ => "1px"
        };
        // CSS has no dash-dot variants; keep their weight and color with a dashed edge.
        string pattern = IsDashDotBorder(style) ? "dashed" : style switch {
            "dotted" => "dotted",
            "dashed" or "mediumdashed" => "dashed",
            "double" => "double",
            _ => "solid"
        };
        AppendCss(css, "border-" + edge, width + " " + pattern + " #" + (ToCssHex(border.ColorArgb) ?? "000000"));
    }

    private static ExcelCellStyleInfo ResolveFallbackMergedBorder(ExcelSheet sheet, ExcelMergeExportRange merge,
        ExcelCellStyleInfo style, IList<HtmlDiagnostic> diagnostics) =>
        style.CopyWithBorder(ExcelMergedCellBorderResolver.Resolve(merge.StartRow, merge.StartColumn, merge.EndRow, merge.EndColumn,
            style.Border, (row, column) => sheet.Cell(row, column).GetStyle().Border, sheet.Name + "!" + merge.A1Range,
            diagnostic => diagnostics.Add(ToHtmlDiagnostic(diagnostic, 1))));

    private static bool IsDashDotBorder(string? style) => style?.Trim().ToLowerInvariant()
        is "dashdot" or "dashdotdot" or "mediumdashdot" or "mediumdashdotdot" or "slantdashdot";

    private static void ReportCssBorderFidelity(ExcelCellBorderInfo? border, int row, int column, CellFidelityDiagnostics diagnostics) {
        if (border == null) return;
        ReportEdge("top", border.Top);
        ReportEdge("right", border.Right);
        ReportEdge("bottom", border.Bottom);
        ReportEdge("left", border.Left);
        if ((border.DiagonalUp || border.DiagonalDown) && border.Diagonal != null
            && !string.IsNullOrWhiteSpace(border.Diagonal.Style)
            && !string.Equals(border.Diagonal.Style, "none", StringComparison.OrdinalIgnoreCase)) {
            diagnostics.Add(HtmlConversionDiagnosticCodes.ContentOmitted,
                "The diagonal cell border is not represented by semantic HTML cell edges.", row, column, "diagonal", OfficeConversionLossKind.Omission);
        }

        void ReportEdge(string edge, ExcelBorderSideInfo? side) {
            if (!IsDashDotBorder(side?.Style)) return;
            diagnostics.Add(HtmlConversionDiagnosticCodes.ContentApproximated,
                "Excel border pattern '" + side!.Style + "' is approximated by a CSS dashed edge.", row, column, edge, OfficeConversionLossKind.Approximation);
        }
    }
}

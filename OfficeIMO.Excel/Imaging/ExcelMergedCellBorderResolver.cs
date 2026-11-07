using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel;

/// <summary>Resolves uniform merged-cell perimeter edges from the existing native style snapshots.</summary>
internal static class ExcelMergedCellBorderResolver {
    /// <summary>Combines uniform outer edges and reports segments that cannot fit one merged-cell edge.</summary>
    /// <remarks>Missing perimeter definitions inherit the origin edge; explicit none overrides it.</remarks>
    internal static ExcelCellBorderInfo? Resolve(
        int firstRow, int firstColumn, int lastRow, int lastColumn,
        ExcelCellBorderInfo? original, Func<int, int, ExcelCellBorderInfo?> borderAt,
        string source, Action<OfficeImageExportDiagnostic> report) {
        ExcelBorderSideInfo? top = Edge("top", firstRow, firstColumn, 0, 1, lastColumn - firstColumn + 1, original?.Top, static border => border?.Top);
        ExcelBorderSideInfo? right = Edge("right", firstRow, lastColumn, 1, 0, lastRow - firstRow + 1, original?.Right, static border => border?.Right);
        ExcelBorderSideInfo? bottom = Edge("bottom", lastRow, firstColumn, 0, 1, lastColumn - firstColumn + 1, original?.Bottom, static border => border?.Bottom);
        ExcelBorderSideInfo? left = Edge("left", firstRow, firstColumn, 1, 0, lastRow - firstRow + 1, original?.Left, static border => border?.Left);
        if (ReferenceEquals(top, original?.Top) && ReferenceEquals(right, original?.Right)
            && ReferenceEquals(bottom, original?.Bottom) && ReferenceEquals(left, original?.Left)) return original;
        return new ExcelCellBorderInfo(left, right, top, bottom, original?.Diagonal,
            original?.DiagonalUp == true, original?.DiagonalDown == true);

        ExcelBorderSideInfo? Edge(string name, int row, int column, int rowStep, int columnStep, int count,
            ExcelBorderSideInfo? fallback, Func<ExcelCellBorderInfo?, ExcelBorderSideInfo?> select) {
            ExcelBorderSideInfo? first = null;
            for (int index = 0; index < count; index++) {
                ExcelBorderSideInfo? current = select(borderAt(row + index * rowStep, column + index * columnStep)) ?? fallback;
                if (index == 0) first = current;
                else if (!SameEdge(first, current)) {
                    report(ExcelImageExportDiagnosticClassifier.Create(OfficeImageExportDiagnosticSeverity.Warning,
                        ExcelImageExportDiagnosticCodes.MergedCellBorderOmitted,
                        "Merged-cell " + name + " border segments differ; the origin edge is retained and differing segments are omitted.",
                        source + ":" + name));
                    return fallback;
                }
            }
            return first;
        }
    }

    private static bool SameEdge(ExcelBorderSideInfo? first, ExcelBorderSideInfo? second) {
        string firstStyle = NormalizeStyle(first?.Style);
        string secondStyle = NormalizeStyle(second?.Style);
        if (!string.Equals(firstStyle, secondStyle, StringComparison.OrdinalIgnoreCase)) return false;
        return string.Equals(firstStyle, "none", StringComparison.OrdinalIgnoreCase)
            || (ExcelRangeImageRenderer.ResolveArgb(first?.ColorArgb) ?? OfficeColor.Black)
                .Equals(ExcelRangeImageRenderer.ResolveArgb(second?.ColorArgb) ?? OfficeColor.Black);
    }

    private static string NormalizeStyle(string? style) => string.IsNullOrWhiteSpace(style) ? "none" : style!.Trim();
}

using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// Resolves percentage tables against Word's compatibility-dependent frame,
    /// after effective edge-cell margins and borders have been resolved.
    /// </summary>
    private static void ApplyNativeTablePercentageWidth(WordTable table, TableLayout layout,
        PdfCore.PdfTableStyle style, double? contentWidth, NativeTableStyleDefaults defaults) {
        if (!contentWidth.HasValue) return;
        W.TableWidth? sourceWidth = table._tableProperties?.TableWidth;
        if (!GetNativeTablePreferredWidth(sourceWidth, contentWidth).HasValue) sourceWidth = defaults.PreferredWidth;
        if (sourceWidth == null || ResolveNativeTablePreferredWidthType(sourceWidth) != W.TableWidthUnitValues.Pct) return;
        double? percent = GetNativeTablePreferredWidthPercent(sourceWidth);
        if (!percent.HasValue) return;

        var (leftBorder, rightBorder) = GetNativeTableOuterBorderInsets(layout, style);
        double referenceWidth = contentWidth.Value - leftBorder - rightBorder;
        bool legacyFrame = !UsesModernNativeWordLayout(table.Document) &&
            !UsesNativeSeparateTableCellBorders(table, style, defaults) && layout.Rows.Count > 0 && layout.Rows[0].Count > 0;
        if (legacyFrame) {
            int firstColumn = GetNativeTableRowStartColumn(layout, 0);
            int lastColumn = GetNativeTableLastCellStartColumn(layout, 0);
            referenceWidth += GetNativeTableCellHorizontalMargin(style, 0, firstColumn, right: false) +
                GetNativeTableCellHorizontalMargin(style, 0, lastColumn, right: true);
        }

        double width = referenceWidth * percent.Value;
        if (width <= 0D) return;
        if (legacyFrame && style.AutoFitColumns) {
            // This is a preferred outer frame, not a cap on unbreakable content.
            style.PreferredWidth = width;
            style.AllowWidthBeyondFrame = true;
        } else if (!style.AutoFitColumns && style.AllowWidthBeyondFrame) {
            style.MaxWidth = width;
        }
    }

    /// <summary>
    /// An unstyled legacy percentage table defaults its text indent to the
    /// first cell margin. An effective table style supplies its own zero or
    /// authored indent, even when Normal Table's other child formatting is ignored.
    /// </summary>
    private static void ApplyNativeLegacyPercentageTableDefaultIndent(WordTable table, TableLayout layout,
        PdfCore.PdfTableStyle style, NativeTableStyleDefaults defaults) {
        if (UsesModernNativeWordLayout(table.Document) ||
            MapNativeTableAlignment(ResolveNativeTableAlignment(table, defaults)) != PdfCore.PdfAlign.Left ||
            GetNativeTableLeftIndent(table._tableProperties?.TableIndentation).HasValue || defaults.LeftIndent.HasValue)
            return;

        W.TableWidth? sourceWidth = table._tableProperties?.TableWidth;
        if (ResolveNativeTablePreferredWidthType(sourceWidth) != W.TableWidthUnitValues.Pct)
            sourceWidth = defaults.PreferredWidth;
        if (sourceWidth == null || ResolveNativeTablePreferredWidthType(sourceWidth) != W.TableWidthUnitValues.Pct ||
            !GetNativeTablePreferredWidthPercent(sourceWidth).HasValue) return;

        NativeStyleLookupCache? styles = GetNativeStyleLookupCache(table.Document);
        string? styleId = styles?.ResolveTableStyleId(GetNativeTableStyleId(table));
        if (styleId != null && styles!.TableStyles.ContainsKey(styleId)) return;
        style.LeftIndent = GetNativeTableCellHorizontalMargin(style, 0,
            GetNativeTableRowStartColumn(layout, 0), right: false);
    }

    private static int GetNativeTableLastCellStartColumn(TableLayout layout, int row) {
        int column = GetNativeTableRowStartColumn(layout, row);
        int lastColumn = column;
        foreach (WordTableCell cell in layout.Rows[row]) {
            if (IsNativeHorizontalMergeContinuation(cell)) continue;
            lastColumn = column;
            column += GetNativeCellColumnSpan(cell);
        }
        return lastColumn;
    }
}

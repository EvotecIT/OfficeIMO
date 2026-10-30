using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Admits only source profiles supported by the shared stable-width contained-flow owner.</summary>
    private static bool CanProjectNativeContainedTable(WordTable table, WordPdfExportOptions? options,
        NativeDocumentDefaults defaults, NativeFontMap fonts, int depth) {
        EnsureNativeTableDepth(depth);
        options?.CancellationToken.ThrowIfCancellationRequested();
        W.Table source = table._table;
        if (depth == 0 && NativeTableStoryHasVariableFrame(source)) return false;
        if (source.GetFirstChild<W.TableProperties>()?.TablePositionProperties != null ||
            source.Descendants<W.TextDirection>().Any(direction =>
                direction.Val?.InnerText is { } directionValue && directionValue != "lrTb") ||
            source.Descendants().Any(element => element is W.SdtBlock or W.SdtRun or W.SdtCell or W.SdtRow) ||
            source.Descendants<W.FootnoteReference>().Any() || source.Descendants<W.EndnoteReference>().Any())
            return false;
        if (source.ChildElements.Any(child => child is not W.TableProperties and not W.TableGrid and not W.TableRow) ||
            source.Elements<W.TableRow>().Any(row => row.ChildElements.Any(child => child is not W.TableRowProperties and not W.TableCell)) ||
            source.Elements<W.TableRow>().SelectMany(row => row.Elements<W.TableCell>()).Any(cell =>
                cell.ChildElements.Any(child => child is not W.TableCellProperties and not W.Paragraph and not W.Table)))
            return false;

        TableLayout layout = TableLayoutCache.GetLayout(table);
        NativeTableStyleDefaults tableDefaults = GetNativeTableStyleDefaults(table, defaults,
            ignoreFallbackTableStyle: options?.PdfOptions?.HasExplicitDefaultTableStyle == true);
        PdfCore.PdfTableStyle style = CreateNativeTableStyle(table, layout.Rows.Count, options, null,
            defaults, tableDefaults, layout, fonts);
        if (style.Position != null || !style.ConsumesVerticalFlow || style.Caption != null || style.ShrinkTextToFit ||
            style.FixedRowHeights?.Any(height => height.HasValue) == true)
            return false;
        foreach (IReadOnlyList<WordTableCell> row in layout.Rows) {
            foreach (WordTableCell cell in row) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (GetNativeCellRowSpan(cell) > 1 ||
                    style.AutoFitColumns && WordTableWidthNative.CanDisableCellWrapping(cell) && !cell.WrapText)
                    return false;
                NativeTableCellEmbeddedContent content = CreateNativeTableCellEmbeddedContent(cell, options,
                    interactiveControls: defaults.RunningContentContext == null);
                if (content.CheckBoxes.Count > 0 || content.FormFields.Count > 0 || content.Images.Count > 0)
                    return false;
                foreach (WordTable nested in EnumerateNativeTableCellElements(cell).OfType<WordTable>())
                    if (!CanProjectNativeContainedTable(nested, options, defaults, fonts, depth + 1)) return false;
            }
        }
        return true;
    }

    /// <summary>
    /// Contained table continuation currently needs one stable physical frame. An
    /// anchored object elsewhere in the same story can leave an active exclusion
    /// when this table is materialized, so decline that story before capturing blocks.
    /// </summary>
    private static bool NativeTableStoryHasVariableFrame(W.Table table) {
        OpenXmlElement story = table.Ancestors().FirstOrDefault(element =>
            element is W.Body or W.Header or W.Footer or W.Footnote or W.Endnote) ?? table;
        return story.Descendants<W.TablePositionProperties>().Any() || story.Descendants<DW.Anchor>().Any();
    }

    /// <summary>
    /// Keeps source percentage widths relative until the shared owner knows the actual
    /// containing cell frame. Body-only legacy text-margin translations do not
    /// apply to this already resolved physical cell frame.
    /// </summary>
    private static void ApplyNativeContainedTablePercentageWidth(WordTable table, TableLayout layout,
        PdfCore.PdfTableStyle style, NativeTableStyleDefaults defaults) {
        W.TableWidth? source = table._tableProperties?.TableWidth;
        if (source == null || ResolveNativeTablePreferredWidthType(source) != W.TableWidthUnitValues.Pct ||
            !GetNativeTablePreferredWidthPercent(source).HasValue) {
            // A valid absolute or automatic source width owns its normal sizing policy.
            if (source != null && (ResolveNativeTablePreferredWidthType(source) == W.TableWidthUnitValues.Auto ||
                GetNativeTablePreferredWidth(source, null).HasValue)) return;
            source = defaults.PreferredWidth;
        }
        if (source == null || ResolveNativeTablePreferredWidthType(source) != W.TableWidthUnitValues.Pct ||
            GetNativeTablePreferredWidthPercent(source) is not { } fraction) return;
        var (leftBorder, rightBorder) = GetNativeTableOuterBorderInsets(layout, style);
        style.FrameWidthAdjustment = -leftBorder - rightBorder;
        // The declared outer width includes the separate-border spacing envelope.
        // This adjustment follows percentage scaling, just as it does for an
        // absolute fixed width in ApplyNativeFixedTableMinimumWidths.
        style.ResolvedWidthAdjustment = style.BorderFrame != null
            ? -2D * style.CellSpacing : 0D;
        if (style.AutoFitColumns) {
            style.PreferredWidth = null;
            style.PreferredWidthFraction = fraction;
            style.PreserveWidth = true;
            style.AutoFitWidthUsesContentMinimum = style.CellSpacing <= 0D;
            List<double?> sourcePreferences = CreateNativeAutoFitColumnPreferences(layout);
            if (sourcePreferences.Count > 0 && sourcePreferences.All(value => value is > 0D)) {
                // Fully authored cell boxes in a percentage frame reserve their
                // physical margins before the stored source preferences shrink.
                // Automatic/null columns keep their ordinary content allocation.
                style.AutoFitColumnPreferredWidthPoints = sourcePreferences;
                style.ColumnWidthWeights = null;
                style.MaxWidth = null;
                style.MaxWidthFraction = fraction;
                style.AutoFitWidthUsesContentMinimum = true;
                style.PreserveAuthoredPreferencesAtWidthLimit = true;
            }
        } else {
            style.MaxWidth = null;
            style.MaxWidthFraction = fraction;
            style.PreserveWidth = true;
            style.AllowWidthBeyondFrame = true;
        }
    }
}

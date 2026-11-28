using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;


public static partial class WordPdfConverterExtensions {
    /// <summary>Admits only source profiles supported by the shared stable-width contained-flow owner.</summary>
    private static bool CanProjectNativeContainedTable(WordTable table, WordPdfExportOptions? options,
        NativeDocumentDefaults defaults, NativeFontMap fonts, int depth) {
        EnsureNativeTableDepth(depth);
        options?.CancellationToken.ThrowIfCancellationRequested();
        OfficeXmlNode source = table.Node();
        if (depth == 0 && NativeTableStoryHasVariableFrame(source)) return false;
        if (source.WChild("tblPr").WChild("tblpPr").Exists ||
            source.DescendantElements(WordXml.W, "textDirection").Any(direction =>
                direction.WVal() is { } directionValue && directionValue != "lrTb") ||
            source.DescendantElements().Any(element => element.IsW("sdt")) ||
            source.HasDescendant(WordXml.W, "footnoteReference") || source.HasDescendant(WordXml.W, "endnoteReference"))
            return false;
        if (source.Elements().Any(child => !child.IsW("tblPr") && !child.IsW("tblGrid") && !child.IsW("tr")) ||
            source.Elements(WordXml.W, "tr").Any(row => row.Elements().Any(child => !child.IsW("trPr") && !child.IsW("tc"))) ||
            source.Elements(WordXml.W, "tr").SelectMany(row => row.Elements(WordXml.W, "tc")).Any(cell =>
                cell.Elements().Any(child => !child.IsW("tcPr") && !child.IsW("p") && !child.IsW("tbl"))))
            return false;

        TableLayout layout = TableLayoutCache.GetLayout(table);
        NativeTableStyleDefaults tableDefaults = GetNativeTableStyleDefaults(table, defaults,
            ignoreFallbackTableStyle: options?.PdfOptions?.HasExplicitDefaultTableStyle == true);
        PdfCore.PdfTableStyle style = CreateNativeTableStyle(table, layout.Rows.Count, options, null,
            defaults, tableDefaults, layout, fonts);
        if (style.Position != null || !style.ConsumesVerticalFlow || style.Caption != null || style.ShrinkTextToFit ||
            style.FixedRowHeights?.ConvertAll(static length => length?.ToPoints())?.Any(height => height.HasValue) == true)
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
    private static bool NativeTableStoryHasVariableFrame(OfficeXmlNode table) {
        OfficeXmlNode story = table.Ancestors().FirstOrDefault(element =>
            element.IsW("body") || element.IsW("hdr") || element.IsW("ftr") || element.IsW("footnote") || element.IsW("endnote"));
        if (story.IsNone) story = table;
        return story.HasDescendant(WordXml.W, "tblpPr") || story.HasDescendant(WordXml.Wp, "anchor");
    }

    /// <summary>
    /// Keeps source percentage widths relative until the shared owner knows the actual
    /// containing cell frame. Body-only legacy text-margin translations do not
    /// apply to this already resolved physical cell frame.
    /// </summary>
    private static void ApplyNativeContainedTablePercentageWidth(WordTable table, TableLayout layout,
        PdfCore.PdfTableStyle style, NativeTableStyleDefaults defaults) {
        OfficeXmlNode source = table.PropertiesNode().WChild("tblW");
        if (source.IsNone || ResolveNativeTablePreferredWidthType(source) != "pct" ||
            !GetNativeTablePreferredWidthPercent(source).HasValue) {
            // A valid absolute or automatic source width owns its normal sizing policy.
            if (source.Exists && (ResolveNativeTablePreferredWidthType(source) == "auto" ||
                GetNativeTablePreferredWidth(source, null).HasValue)) return;
            source = defaults.PreferredWidth;
        }
        if (source.IsNone || ResolveNativeTablePreferredWidthType(source) != "pct" ||
            GetNativeTablePreferredWidthPercent(source) is not { } fraction) return;
        var (leftBorder, rightBorder) = GetNativeTableOuterBorderInsets(layout, style);
        style.FrameWidthAdjustment = -leftBorder - rightBorder;
        // The declared outer width includes the separate-border spacing envelope.
        // This adjustment follows percentage scaling, just as it does for an
        // absolute fixed width in ApplyNativeFixedTableMinimumWidths.
        style.ResolvedWidthAdjustment = style.BorderFrame != null
            ? -2D * style.CellSpacing.ToPoints() : 0D;
        if (style.AutoFitColumns) {
            style.PreferredWidth = null;
            style.PreferredWidthFraction = fraction;
            style.PreserveWidth = true;
            style.AutoFitWidthUsesContentMinimum = style.CellSpacing.ToPoints() <= 0D;
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

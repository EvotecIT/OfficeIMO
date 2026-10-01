namespace OfficeIMO.Rtf.Writing;

internal sealed partial class RtfWriteContext {
    private sealed class MaterializationCounts {
        internal int Paragraphs;
        internal int Runs;
    }

    internal void TrackMaterialization(RtfConversionReport? report) {
        if (FormattingDocument != null && report != null) Counts = new MaterializationCounts();
    }

    internal void ReportMaterialization(RtfConversionReport? report) {
        if (Counts == null || Counts.Paragraphs + Counts.Runs == 0 || report == null) return;
        report.Add(RtfConversionSeverity.Warning, "RtfNormalizationFormattingInheritanceMaterialized",
            "Effective paragraph and character formatting is written as direct body controls. Appearance is retained; reopening this output loses the distinction between inherited and direct formatting.",
            RtfConversionAction.Flattened, sourcePath: "Document", feature: "FormattingInheritance",
            count: Counts.Paragraphs + Counts.Runs,
            detail: "Paragraphs=" + Counts.Paragraphs.ToString(CultureInfo.InvariantCulture) + ";Runs=" + Counts.Runs.ToString(CultureInfo.InvariantCulture));
    }

    private static bool HasMaterializedRunFormatting(RtfRun source, RtfRun resolved) =>
        source.DirectBold != resolved.DirectBold || source.DirectItalic != resolved.DirectItalic ||
        source.DirectHidden != resolved.DirectHidden || source.DirectUnderlineStyle != resolved.DirectUnderlineStyle ||
        source.FontSize != resolved.FontSize || source.FontId != resolved.FontId ||
        source.ForegroundColorIndex != resolved.ForegroundColorIndex || source.HighlightColorIndex != resolved.HighlightColorIndex;

    private static bool HasMaterializedParagraphFormatting(RtfParagraph source, RtfParagraph resolved) =>
        source.ListId != resolved.ListId || source.ListLevel != resolved.ListLevel ||
        source.DirectAlignment != resolved.DirectAlignment || source.Direction != resolved.Direction ||
        source.DirectPageBreakBefore != resolved.DirectPageBreakBefore || source.DirectKeepWithNext != resolved.DirectKeepWithNext ||
        source.DirectKeepLinesTogether != resolved.DirectKeepLinesTogether || source.DirectSuppressLineNumbers != resolved.DirectSuppressLineNumbers ||
        source.LeftIndentTwips != resolved.LeftIndentTwips || source.RightIndentTwips != resolved.RightIndentTwips ||
        source.FirstLineIndentTwips != resolved.FirstLineIndentTwips || source.SpaceBeforeTwips != resolved.SpaceBeforeTwips ||
        source.SpaceAfterTwips != resolved.SpaceAfterTwips || source.SpaceBeforeAuto != resolved.SpaceBeforeAuto ||
        source.SpaceAfterAuto != resolved.SpaceAfterAuto || source.LineSpacingTwips != resolved.LineSpacingTwips ||
        source.LineSpacingMultiple != resolved.LineSpacingMultiple || source.BackgroundColorIndex != resolved.BackgroundColorIndex ||
        source.ShadingForegroundColorIndex != resolved.ShadingForegroundColorIndex || source.ShadingPatternPercent != resolved.ShadingPatternPercent ||
        source.ShadingPattern != resolved.ShadingPattern || source.AutoHyphenation != resolved.AutoHyphenation ||
        source.ContextualSpacing != resolved.ContextualSpacing || source.AdjustRightIndent != resolved.AdjustRightIndent ||
        source.SnapToLineGrid != resolved.SnapToLineGrid || source.WidowControl != resolved.WidowControl || source.OutlineLevel != resolved.OutlineLevel ||
        (source.TabStops.Count == 0 && resolved.TabStops.Count > 0) ||
        (!source.TopBorder.HasAnyValue && resolved.TopBorder.HasAnyValue) || (!source.LeftBorder.HasAnyValue && resolved.LeftBorder.HasAnyValue) ||
        (!source.BottomBorder.HasAnyValue && resolved.BottomBorder.HasAnyValue) || (!source.RightBorder.HasAnyValue && resolved.RightBorder.HasAnyValue) ||
        (!source.Frame.HasAnyValue && resolved.Frame.HasAnyValue) || (!source.LegacyNumbering.HasAnyValue && resolved.LegacyNumbering.HasAnyValue);
}

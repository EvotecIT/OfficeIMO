namespace OfficeIMO.Rtf;

/// <content>Resolves authored style references through the shared semantic model.</content>
public sealed partial class RtfDocument {
    /// <summary>Returns an independent paragraph copy with inherited paragraph formatting applied. Missing bases and cycles stop the style chain at the last reachable entry.</summary>
    public RtfParagraph ResolveParagraphFormatting(RtfParagraph paragraph) {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        var result = new RtfCloneContext().Clone(paragraph)!;
        return ApplyParagraphListFormatting(ApplyParagraphInheritance(result));
    }

    internal RtfParagraph GetParagraphFormatting(RtfParagraph paragraph) => ApplyParagraphListFormatting(ApplyParagraphInheritance(paragraph.CopyFormattingView()));

    private RtfParagraph ApplyParagraphInheritance(RtfParagraph result) {
        foreach (RtfStyle style in GetFormattingStyleChain(result.StyleId ?? 0, RtfStyleKind.Paragraph).Reverse()) {
            result.ListId ??= style.ListId;
            result.ListLevel ??= style.ListLevel;
            result.DirectAlignment ??= style.ParagraphAlignment;
            result.DirectPageBreakBefore ??= style.PageBreakBefore;
            result.DirectKeepWithNext ??= style.KeepWithNext;
            result.DirectKeepLinesTogether ??= style.KeepLinesTogether;
            result.DirectSuppressLineNumbers ??= style.SuppressLineNumbers;
            result.Direction ??= style.ParagraphDirection;
            result.LeftIndentTwips ??= style.LeftIndentTwips;
            result.RightIndentTwips ??= style.RightIndentTwips;
            result.FirstLineIndentTwips ??= style.FirstLineIndentTwips;
            result.SpaceBeforeTwips ??= style.SpaceBeforeTwips;
            result.SpaceAfterTwips ??= style.SpaceAfterTwips;
            result.SpaceBeforeAuto ??= style.SpaceBeforeAuto;
            result.SpaceAfterAuto ??= style.SpaceAfterAuto;
            result.LineSpacingTwips ??= style.LineSpacingTwips;
            result.LineSpacingMultiple ??= style.LineSpacingMultiple;
            result.BackgroundColorIndex ??= style.BackgroundColorIndex;
            result.ShadingForegroundColorIndex ??= style.ShadingForegroundColorIndex;
            result.ShadingPatternPercent ??= style.ShadingPatternPercent;
            result.AutoHyphenation ??= style.AutoHyphenation;
            result.ContextualSpacing ??= style.ContextualSpacing;
            result.AdjustRightIndent ??= style.AdjustRightIndent;
            result.SnapToLineGrid ??= style.SnapToLineGrid;
            result.WidowControl ??= style.WidowControl;
            result.OutlineLevel ??= style.OutlineLevel;
            if (result.TabStops.Count == 0 && style.TabStops.Count > 0) result.ReplaceTabStops(style.TabStops);
            if (result.ShadingPattern == RtfShadingPattern.None) result.ShadingPattern = style.ShadingPattern;
            InheritBorder(result.TopBorder, style.TopBorder);
            InheritBorder(result.LeftBorder, style.LeftBorder);
            InheritBorder(result.BottomBorder, style.BottomBorder);
            InheritBorder(result.RightBorder, style.RightBorder);
            if (!result.Frame.HasAnyValue) result.Frame.CopyFrom(style.Frame);
            if (!result.LegacyNumbering.HasAnyValue) result.LegacyNumbering.CopyFrom(style.LegacyNumbering);
        }
        return result;
    }

    private RtfParagraph ApplyParagraphListFormatting(RtfParagraph paragraph) {
        if (paragraph.ListId == 0) {
            paragraph.ListKind = RtfListKind.None;
            paragraph.ListDefinitionId = null;
            return paragraph;
        }
        RtfListFormatting? list = ResolveListFormattingCore(paragraph);
        if (list != null) {
            paragraph.ListKind = list.Level.Kind;
            paragraph.ListId = list.InstanceId;
            paragraph.ListDefinitionId = list.DefinitionId;
            paragraph.ListLevel = list.LevelIndex;
            paragraph.LeftIndentTwips ??= list.Level.LeftIndentTwips;
            paragraph.FirstLineIndentTwips ??= list.Level.FirstLineIndentTwips;
        }
        return paragraph;
    }

    /// <summary>Returns an independent run copy with paragraph and character styles applied beneath direct formatting.</summary>
    public RtfRun ResolveRunFormatting(RtfParagraph paragraph, RtfRun run) {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        if (run == null) throw new ArgumentNullException(nameof(run));
        var result = new RtfCloneContext().Clone(run)!;
        return ApplyRunInheritance(paragraph, result);
    }

    internal RtfRun GetRunFormatting(RtfParagraph paragraph, RtfRun run) => ApplyRunInheritance(paragraph, run.CopyFormattingView());

    private RtfRun ApplyRunInheritance(RtfParagraph paragraph, RtfRun result) {
        IEnumerable<RtfStyle> styles = GetFormattingStyleChain(result.StyleId, RtfStyleKind.Character).Reverse();
        if (!result.UseDefaultCharacterFormatting) styles = styles.Concat(GetFormattingStyleChain(paragraph.StyleId ?? 0, RtfStyleKind.Paragraph).Reverse());
        foreach (RtfStyle style in styles) {
            result.DirectBold ??= style.Bold;
            result.DirectItalic ??= style.Italic;
            result.DirectUnderlineStyle ??= style.UnderlineStyle;
            result.FontSize ??= style.FontSize;
            result.FontId ??= style.FontId;
            result.ForegroundColorIndex ??= style.ForegroundColorIndex;
            result.HighlightColorIndex ??= style.HighlightColorIndex;
        }
        if (result.UseDefaultCharacterFormatting) {
            result.FontSize ??= 12;
            result.FontId ??= Settings.DefaultFontId ?? Fonts.FirstOrDefault()?.Id;
            result.ForegroundColorIndex ??= 0;
            result.HighlightColorIndex ??= 0;
        }
        return result;
    }

    private IReadOnlyList<RtfStyle> GetFormattingStyleChain(int? id, RtfStyleKind kind) {
        var chain = new Stack<RtfStyle>();
        var visited = new HashSet<int>();
        while (id.HasValue && visited.Add(id.Value)) {
            RtfStyle? current = Styles.LastOrDefault(style => style.Id == id.Value && style.Kind == kind);
            if (current == null) break;
            chain.Push(current);
            id = current.BasedOnStyleId;
        }
        return chain.ToArray();
    }

    private static void InheritBorder(RtfParagraphBorder target, RtfParagraphBorder source) {
        if (target.HasAnyValue || !source.HasAnyValue) return;
        target.DirectStyle = source.DirectStyle;
        target.Width = source.Width;
        target.ColorIndex = source.ColorIndex;
    }
}

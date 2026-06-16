using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyParagraphLayout(WordParagraph source, RtfParagraph destination) {
        destination.PageBreakBefore = source.PageBreakBefore;
        destination.KeepWithNext = source.KeepWithNext;
        destination.KeepLinesTogether = source.KeepLinesTogether;
        destination.SuppressLineNumbers = source._paragraphProperties?.SuppressLineNumbers != null;
        destination.AutoHyphenation = source._paragraphProperties?.SuppressAutoHyphens == null ? null : false;
        destination.ContextualSpacing = GetContextualSpacing(source);
        destination.AdjustRightIndent = GetAdjustRightIndent(source);
        destination.SnapToLineGrid = GetSnapToLineGrid(source);
        destination.WidowControl = GetWidowControl(source);
        Int32Value? outlineLevel = source._paragraphProperties?.OutlineLevel?.Val;
        destination.OutlineLevel = outlineLevel?.Value;
        destination.SpaceBeforeTwips = source.LineSpacingBefore;
        destination.SpaceAfterTwips = source.LineSpacingAfter;
        SpacingBetweenLines? spacing = source._paragraphProperties?.SpacingBetweenLines;
        destination.SpaceBeforeAuto = spacing?.BeforeAutoSpacing?.Value;
        destination.SpaceAfterAuto = spacing?.AfterAutoSpacing?.Value;
        destination.LineSpacingTwips = source.LineSpacing;
        destination.LineSpacingMultiple = ToRtfLineSpacingMultiple(source.LineSpacingRule);
        destination.Direction = source.BiDi ? RtfTextDirection.RightToLeft : null;
    }

    private static void ApplyParagraphLayout(WordParagraph destination, RtfParagraph source) {
        destination.PageBreakBefore = source.PageBreakBefore;
        destination.KeepWithNext = source.KeepWithNext;
        destination.KeepLinesTogether = source.KeepLinesTogether;
        if (source.SuppressLineNumbers) {
            ParagraphProperties properties = destination._paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.SuppressLineNumbers = new SuppressLineNumbers();
        }

        if (source.AutoHyphenation == false) {
            ParagraphProperties properties = destination._paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.SuppressAutoHyphens = new SuppressAutoHyphens();
        }

        if (source.ContextualSpacing.HasValue) {
            ParagraphProperties properties = destination._paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.ContextualSpacing = new ContextualSpacing { Val = source.ContextualSpacing.Value };
        }

        if (source.AdjustRightIndent.HasValue) {
            ParagraphProperties properties = destination._paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.AdjustRightIndent = new AdjustRightIndent { Val = source.AdjustRightIndent.Value };
        }

        if (source.SnapToLineGrid.HasValue) {
            ParagraphProperties properties = destination._paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.SnapToGrid = new SnapToGrid { Val = source.SnapToLineGrid.Value };
        }

        if (source.Direction.HasValue) {
            destination.BiDi = source.Direction.Value == RtfTextDirection.RightToLeft;
        }

        if (source.SpaceBeforeTwips.HasValue) {
            destination.LineSpacingBefore = source.SpaceBeforeTwips.Value;
        }

        if (source.SpaceAfterTwips.HasValue) {
            destination.LineSpacingAfter = source.SpaceAfterTwips.Value;
        }

        if (source.SpaceBeforeAuto.HasValue || source.SpaceAfterAuto.HasValue) {
            ParagraphProperties properties = destination._paragraph.ParagraphProperties ??= new ParagraphProperties();
            SpacingBetweenLines spacing = properties.SpacingBetweenLines ??= new SpacingBetweenLines();
            if (source.SpaceBeforeAuto.HasValue) {
                spacing.BeforeAutoSpacing = source.SpaceBeforeAuto.Value;
            }

            if (source.SpaceAfterAuto.HasValue) {
                spacing.AfterAutoSpacing = source.SpaceAfterAuto.Value;
            }
        }

        if (source.LineSpacingTwips.HasValue) {
            destination.LineSpacing = source.LineSpacingTwips.Value;
        }

        LineSpacingRuleValues? lineSpacingRule = ToWordLineSpacingRule(source.LineSpacingMultiple);
        if (lineSpacingRule.HasValue) {
            destination.LineSpacingRule = lineSpacingRule.Value;
        }

        if (source.WidowControl.HasValue) {
            ParagraphProperties properties = destination._paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.WidowControl = new WidowControl { Val = source.WidowControl.Value };
        }

        if (source.OutlineLevel.HasValue) {
            ParagraphProperties properties = destination._paragraph.ParagraphProperties ??= new ParagraphProperties();
            properties.OutlineLevel = new OutlineLevel { Val = source.OutlineLevel.Value };
        }
    }

    private static bool? GetWidowControl(WordParagraph source) {
        WidowControl? widowControl = source._paragraphProperties?.WidowControl;
        if (widowControl == null) {
            return null;
        }

        return widowControl.Val?.Value ?? true;
    }

    private static bool? GetContextualSpacing(WordParagraph source) {
        ContextualSpacing? contextualSpacing = source._paragraphProperties?.ContextualSpacing;
        if (contextualSpacing == null) {
            return null;
        }

        return contextualSpacing.Val?.Value ?? true;
    }

    private static bool? GetAdjustRightIndent(WordParagraph source) {
        AdjustRightIndent? adjustRightIndent = source._paragraphProperties?.AdjustRightIndent;
        if (adjustRightIndent == null) {
            return null;
        }

        return adjustRightIndent.Val?.Value ?? true;
    }

    private static bool? GetSnapToLineGrid(WordParagraph source) {
        SnapToGrid? snapToGrid = source._paragraphProperties?.SnapToGrid;
        if (snapToGrid == null) {
            return null;
        }

        return snapToGrid.Val?.Value ?? true;
    }

    private static bool? ToRtfLineSpacingMultiple(LineSpacingRuleValues? value) {
        if (!value.HasValue) {
            return null;
        }

        return value.Value == LineSpacingRuleValues.Auto;
    }

    private static LineSpacingRuleValues? ToWordLineSpacingRule(bool? multiple) {
        if (!multiple.HasValue) {
            return null;
        }

        return multiple.Value ? LineSpacingRuleValues.Auto : LineSpacingRuleValues.Exact;
    }
}

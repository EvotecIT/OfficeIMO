using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyParagraphLayout(WordParagraph source, RtfParagraph destination) {
        destination.DirectAlignment = ToRtfTextAlignment(source._paragraphProperties?.Justification?.Val?.Value);
        destination.LeftIndentTwips = checked((int?)(source.IndentationBefore?.ToTwips()));
        destination.RightIndentTwips = checked((int?)(source.IndentationAfter?.ToTwips()));
        destination.FirstLineIndentTwips = checked((int?)(source.IndentationFirstLine?.ToTwips())) ?? Negate(checked((int?)(source.IndentationHanging?.ToTwips())));
        destination.DirectPageBreakBefore = ReadToggle(source._paragraphProperties?.PageBreakBefore);
        destination.DirectKeepWithNext = ReadToggle(source._paragraphProperties?.KeepNext);
        destination.DirectKeepLinesTogether = ReadToggle(source._paragraphProperties?.KeepLines);
        destination.DirectSuppressLineNumbers = ReadToggle(source._paragraphProperties?.SuppressLineNumbers);
        destination.AutoHyphenation = source._paragraphProperties?.SuppressAutoHyphens == null ? null : false;
        destination.ContextualSpacing = GetContextualSpacing(source);
        destination.AdjustRightIndent = GetAdjustRightIndent(source);
        destination.SnapToLineGrid = GetSnapToLineGrid(source);
        destination.WidowControl = GetWidowControl(source);
        Int32Value? outlineLevel = source._paragraphProperties?.OutlineLevel?.Val;
        destination.OutlineLevel = outlineLevel?.Value;
        destination.SpaceBeforeTwips = checked((int?)(source.LineSpacingBefore?.ToTwips()));
        destination.SpaceAfterTwips = checked((int?)(source.LineSpacingAfter?.ToTwips()));
        SpacingBetweenLines? spacing = source._paragraphProperties?.SpacingBetweenLines;
        destination.SpaceBeforeAuto = spacing?.BeforeAutoSpacing?.Value;
        destination.SpaceAfterAuto = spacing?.AfterAutoSpacing?.Value;
        destination.LineSpacingTwips = ToRtfLineSpacingValue(source.NativeLineSpacingValue, spacing?.LineRule?.Value);
        destination.LineSpacingMultiple = ToRtfLineSpacingMultiple(source.NativeLineSpacingValue, spacing?.LineRule?.Value);
        destination.Direction = source.BiDi ? RtfTextDirection.RightToLeft : null;
    }

    private static void ApplyParagraphLayout(WordParagraph destination, RtfParagraph source) {
        ParagraphProperties flow = destination._paragraph.ParagraphProperties ??= new ParagraphProperties();
        flow.PageBreakBefore = source.DirectPageBreakBefore.HasValue ? new PageBreakBefore { Val = source.PageBreakBefore } : null;
        flow.KeepNext = source.DirectKeepWithNext.HasValue ? new KeepNext { Val = source.KeepWithNext } : null;
        flow.KeepLines = source.DirectKeepLinesTogether.HasValue ? new KeepLines { Val = source.KeepLinesTogether } : null;
        flow.SuppressLineNumbers = source.DirectSuppressLineNumbers.HasValue ? new SuppressLineNumbers { Val = source.SuppressLineNumbers } : null;
        flow.Justification = source.DirectAlignment.HasValue ? new Justification { Val = ToWordTextAlignment(source.DirectAlignment.Value) } : null;

        if (source.LeftIndentTwips.HasValue || source.RightIndentTwips.HasValue || source.FirstLineIndentTwips.HasValue) {
            flow.Indentation = new Indentation {
                Left = FormatInt(source.LeftIndentTwips),
                Right = FormatInt(source.RightIndentTwips),
                FirstLine = source.FirstLineIndentTwips >= 0 ? FormatInt(source.FirstLineIndentTwips) : null,
                Hanging = source.FirstLineIndentTwips < 0 ? FormatInt(-source.FirstLineIndentTwips) : null
            };
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
            destination.LineSpacingBefore= ((source.SpaceBeforeTwips.Value) is { } wordLength4961 ? global::OfficeIMO.OfficeLength.Points((wordLength4961) / 20D) : (global::OfficeIMO.OfficeLength?)null);
        }

        if (source.SpaceAfterTwips.HasValue) {
            destination.LineSpacingAfter= ((source.SpaceAfterTwips.Value) is { } wordLength5093 ? global::OfficeIMO.OfficeLength.Points((wordLength5093) / 20D) : (global::OfficeIMO.OfficeLength?)null);
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

        int? lineSpacing = ToWordLineSpacingValue(source.LineSpacingTwips, source.LineSpacingMultiple);
        if (lineSpacing.HasValue) destination.NativeLineSpacingValue = lineSpacing.Value;
        LineSpacingRuleValues? lineSpacingRule = ToWordLineSpacingRule(source.LineSpacingTwips, source.LineSpacingMultiple);
        if (lineSpacingRule.HasValue) {
            destination.NativeLineSpacingRule = lineSpacingRule.Value.ToOfficeEnum();
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

}

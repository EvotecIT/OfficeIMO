using OfficeIMO.Word;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyParagraphLayout(WordParagraph source, RtfParagraph destination) {
        OfficeXmlNode properties = source.ParagraphPropertiesNode();
        destination.DirectAlignment = ToRtfTextAlignment(properties.ModelChild("jc").WVal());
        destination.LeftIndentTwips = checked((int?)(source.IndentationBefore?.ToTwips()));
        destination.RightIndentTwips = checked((int?)(source.IndentationAfter?.ToTwips()));
        destination.FirstLineIndentTwips = checked((int?)(source.IndentationFirstLine?.ToTwips())) ?? Negate(checked((int?)(source.IndentationHanging?.ToTwips())));
        destination.DirectPageBreakBefore = ReadToggle(properties.ModelChild("pageBreakBefore"));
        destination.DirectKeepWithNext = ReadToggle(properties.ModelChild("keepNext"));
        destination.DirectKeepLinesTogether = ReadToggle(properties.ModelChild("keepLines"));
        destination.DirectSuppressLineNumbers = ReadToggle(properties.ModelChild("suppressLineNumbers"));
        destination.AutoHyphenation = properties.ModelChild("suppressAutoHyphens").IsNone ? null : false;
        destination.ContextualSpacing = ReadToggle(properties.ModelChild("contextualSpacing"));
        destination.AdjustRightIndent = ReadToggle(properties.ModelChild("adjustRightInd"));
        destination.SnapToLineGrid = ReadToggle(properties.ModelChild("snapToGrid"));
        destination.WidowControl = ReadToggle(properties.ModelChild("widowControl"));
        destination.OutlineLevel = WordXml.ParseInt32(properties.ModelChild("outlineLvl").WVal());
        destination.SpaceBeforeTwips = checked((int?)(source.LineSpacingBefore?.ToTwips()));
        destination.SpaceAfterTwips = checked((int?)(source.LineSpacingAfter?.ToTwips()));
        OfficeXmlNode spacing = properties.ModelChild("spacing");
        destination.SpaceBeforeAuto = WordXml.ParseOnOff(spacing.WAttr("beforeAutospacing"));
        destination.SpaceAfterAuto = WordXml.ParseOnOff(spacing.WAttr("afterAutospacing"));
        string? lineRule = spacing.WAttr("lineRule");
        destination.LineSpacingTwips = ToRtfLineSpacingValue(source.NativeLineSpacingValue, lineRule);
        destination.LineSpacingMultiple = ToRtfLineSpacingMultiple(source.NativeLineSpacingValue, lineRule);
        destination.Direction = source.BiDi ? RtfTextDirection.RightToLeft : null;
    }

    private static void ApplyParagraphLayout(WordParagraph destination, RtfParagraph source) {
        OfficeXmlNode flow = GetOrAddParagraphProperties(destination);
        SetParagraphToggle(flow, "pageBreakBefore", source.DirectPageBreakBefore.HasValue ? source.PageBreakBefore : null);
        SetParagraphToggle(flow, "keepNext", source.DirectKeepWithNext.HasValue ? source.KeepWithNext : null);
        SetParagraphToggle(flow, "keepLines", source.DirectKeepLinesTogether.HasValue ? source.KeepLinesTogether : null);
        SetParagraphToggle(flow, "suppressLineNumbers", source.DirectSuppressLineNumbers.HasValue ? source.SuppressLineNumbers : null);
        if (source.DirectAlignment.HasValue) {
            WordXmlEdits.ReplaceChild(flow, "jc", WPPrView.Order, ToWordTextAlignment(source.DirectAlignment.Value));
        } else {
            WordXmlEdits.RemoveChild(flow, "jc");
        }

        if (source.LeftIndentTwips.HasValue || source.RightIndentTwips.HasValue || source.FirstLineIndentTwips.HasValue) {
            OfficeXmlNode indentation = WordXmlEdits.ReplaceChild(flow, "ind", WPPrView.Order);
            indentation.SetAttribute(WordXml.W, "left", FormatInt(source.LeftIndentTwips));
            indentation.SetAttribute(WordXml.W, "right", FormatInt(source.RightIndentTwips));
            indentation.SetAttribute(WordXml.W, "hanging", source.FirstLineIndentTwips < 0 ? FormatInt(-source.FirstLineIndentTwips) : null);
            indentation.SetAttribute(WordXml.W, "firstLine", source.FirstLineIndentTwips >= 0 ? FormatInt(source.FirstLineIndentTwips) : null);
        }

        if (source.AutoHyphenation == false) {
            WordXmlEdits.ReplaceChild(GetOrAddParagraphProperties(destination), "suppressAutoHyphens", WPPrView.Order);
        }

        if (source.ContextualSpacing.HasValue) {
            SetParagraphToggle(GetOrAddParagraphProperties(destination), "contextualSpacing", source.ContextualSpacing.Value);
        }

        if (source.AdjustRightIndent.HasValue) {
            SetParagraphToggle(GetOrAddParagraphProperties(destination), "adjustRightInd", source.AdjustRightIndent.Value);
        }

        if (source.SnapToLineGrid.HasValue) {
            SetParagraphToggle(GetOrAddParagraphProperties(destination), "snapToGrid", source.SnapToLineGrid.Value);
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
            OfficeXmlNode spacing = GetOrAddParagraphProperties(destination).GetOrAddChild(WordXml.W, "spacing", WPPrView.Order);
            if (source.SpaceBeforeAuto.HasValue) {
                spacing.SetAttribute(WordXml.W, "beforeAutospacing", FormatOnOff(source.SpaceBeforeAuto.Value));
            }

            if (source.SpaceAfterAuto.HasValue) {
                spacing.SetAttribute(WordXml.W, "afterAutospacing", FormatOnOff(source.SpaceAfterAuto.Value));
            }
        }

        int? lineSpacing = ToWordLineSpacingValue(source.LineSpacingTwips, source.LineSpacingMultiple);
        if (lineSpacing.HasValue) destination.NativeLineSpacingValue = lineSpacing.Value;
        string? lineSpacingRule = ToWordLineSpacingRule(source.LineSpacingTwips, source.LineSpacingMultiple);
        if (lineSpacingRule != null && WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(lineSpacingRule, out WordLineSpacingRule rule)) {
            destination.NativeLineSpacingRule = rule;
        }

        if (source.WidowControl.HasValue) {
            SetParagraphToggle(GetOrAddParagraphProperties(destination), "widowControl", source.WidowControl.Value);
        }

        if (source.OutlineLevel.HasValue) {
            WordXmlEdits.ReplaceChild(GetOrAddParagraphProperties(destination), "outlineLvl", WPPrView.Order, FormatInt(source.OutlineLevel.Value)!);
        }
    }

    /// <summary>Puts an on/off property with an explicit value in place, as assigning a typed SDK property does; null removes it.</summary>
    private static void SetParagraphToggle(OfficeXmlNode properties, string localName, bool? value) {
        if (value.HasValue) {
            WordXmlEdits.ReplaceChild(properties, localName, WPPrView.Order, FormatOnOff(value.Value));
        } else {
            WordXmlEdits.RemoveChild(properties, localName);
        }
    }
}

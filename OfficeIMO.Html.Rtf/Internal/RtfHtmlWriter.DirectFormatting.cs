namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendParagraphDirectFormatting(StringBuilder builder, RtfParagraph source) {
        var values = new Dictionary<string, string> { ["version"] = "1" };
        AddNullableInt(values, "ListId", source.ListId);
        AddNullableInt(values, "ListLevel", source.ListLevel);
        AddEnum(values, "DirectAlignment", source.DirectAlignment);
        AddNullableBool(values, "DirectPageBreakBefore", source.DirectPageBreakBefore);
        AddNullableBool(values, "DirectKeepWithNext", source.DirectKeepWithNext);
        AddNullableBool(values, "DirectKeepLinesTogether", source.DirectKeepLinesTogether);
        AddNullableBool(values, "DirectSuppressLineNumbers", source.DirectSuppressLineNumbers);
        AddEnum(values, "Direction", source.Direction);
        AddNullableInt(values, "LeftIndentTwips", source.LeftIndentTwips);
        AddNullableInt(values, "RightIndentTwips", source.RightIndentTwips);
        AddNullableInt(values, "FirstLineIndentTwips", source.FirstLineIndentTwips);
        AddNullableInt(values, "SpaceBeforeTwips", source.SpaceBeforeTwips);
        AddNullableInt(values, "SpaceAfterTwips", source.SpaceAfterTwips);
        AddNullableInt(values, "LineSpacingTwips", source.LineSpacingTwips);
        AddNullableInt(values, "BackgroundColorIndex", source.BackgroundColorIndex);
        AddNullableInt(values, "ShadingForegroundColorIndex", source.ShadingForegroundColorIndex);
        AddNullableInt(values, "ShadingPatternPercent", source.ShadingPatternPercent);
        AddNullableInt(values, "OutlineLevel", source.OutlineLevel);
        AddNullableBool(values, "SpaceBeforeAuto", source.SpaceBeforeAuto);
        AddNullableBool(values, "SpaceAfterAuto", source.SpaceAfterAuto);
        AddNullableBool(values, "LineSpacingMultiple", source.LineSpacingMultiple);
        AddNullableBool(values, "AutoHyphenation", source.AutoHyphenation);
        AddNullableBool(values, "ContextualSpacing", source.ContextualSpacing);
        AddNullableBool(values, "AdjustRightIndent", source.AdjustRightIndent);
        AddNullableBool(values, "SnapToLineGrid", source.SnapToLineGrid);
        AddNullableBool(values, "WidowControl", source.WidowControl);
        AddBorder(values, "border.top", source.TopBorder);
        AddBorder(values, "border.left", source.LeftBorder);
        AddBorder(values, "border.bottom", source.BottomBorder);
        AddBorder(values, "border.right", source.RightBorder);
        builder.Append(" data-officeimo-rtf-direct-paragraph=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append('"');
    }

    private static void AppendRunDirectFormatting(StringBuilder builder, RtfRun source) {
        var values = new Dictionary<string, string> { ["version"] = "1" };
        AddBool(values, "plain", source.UseDefaultCharacterFormatting);
        AddNullableBool(values, "DirectBold", source.DirectBold);
        AddNullableBool(values, "DirectItalic", source.DirectItalic);
        AddEnum(values, "DirectUnderlineStyle", source.DirectUnderlineStyle);
        AddNullableDouble(values, "FontSize", source.FontSize);
        AddNullableInt(values, "FontId", source.FontId);
        AddNullableInt(values, "ForegroundColorIndex", source.ForegroundColorIndex);
        AddNullableInt(values, "HighlightColorIndex", source.HighlightColorIndex);
        builder.Append(" data-officeimo-rtf-direct-run=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append('"');
    }

}

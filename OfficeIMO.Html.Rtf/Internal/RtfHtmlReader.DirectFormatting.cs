namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void RestoreParagraphDirectFormatting(IElement token) {
            var values = RtfHtmlMetadataCodec.Decode(GetAttribute(token, "data-officeimo-rtf-direct-paragraph"));
            if (!values.TryGetValue("version", out string? version) || version != "1") return;
            RtfParagraph source = EnsureParagraph();
            source.ListId = ReadInt(values, "ListId");
            source.ListLevel = ReadInt(values, "ListLevel");
            source.DirectAlignment = ReadEnum<RtfTextAlignment>(values, "DirectAlignment");
            source.DirectPageBreakBefore = ReadBool(values, "DirectPageBreakBefore");
            source.DirectKeepWithNext = ReadBool(values, "DirectKeepWithNext");
            source.DirectKeepLinesTogether = ReadBool(values, "DirectKeepLinesTogether");
            source.DirectSuppressLineNumbers = ReadBool(values, "DirectSuppressLineNumbers");
            source.Direction = ReadEnum<RtfTextDirection>(values, "Direction");
            source.LeftIndentTwips = ReadInt(values, "LeftIndentTwips");
            source.RightIndentTwips = ReadInt(values, "RightIndentTwips");
            source.FirstLineIndentTwips = ReadInt(values, "FirstLineIndentTwips");
            source.SpaceBeforeTwips = ReadInt(values, "SpaceBeforeTwips");
            source.SpaceAfterTwips = ReadInt(values, "SpaceAfterTwips");
            source.LineSpacingTwips = ReadInt(values, "LineSpacingTwips");
            source.BackgroundColorIndex = ReadInt(values, "BackgroundColorIndex");
            source.ShadingForegroundColorIndex = ReadInt(values, "ShadingForegroundColorIndex");
            source.ShadingPatternPercent = ReadInt(values, "ShadingPatternPercent");
            source.OutlineLevel = ReadInt(values, "OutlineLevel");
            source.SpaceBeforeAuto = ReadBool(values, "SpaceBeforeAuto");
            source.SpaceAfterAuto = ReadBool(values, "SpaceAfterAuto");
            source.LineSpacingMultiple = ReadBool(values, "LineSpacingMultiple");
            source.AutoHyphenation = ReadBool(values, "AutoHyphenation");
            source.ContextualSpacing = ReadBool(values, "ContextualSpacing");
            source.AdjustRightIndent = ReadBool(values, "AdjustRightIndent");
            source.SnapToLineGrid = ReadBool(values, "SnapToLineGrid");
            source.WidowControl = ReadBool(values, "WidowControl");
            ApplyBorder(values, "border.top", source.TopBorder);
            ApplyBorder(values, "border.left", source.LeftBorder);
            ApplyBorder(values, "border.bottom", source.BottomBorder);
            ApplyBorder(values, "border.right", source.RightBorder);
        }

        private void RestoreRunDirectFormatting(RtfRun source) {
            Dictionary<string, string>? values = _styles.FirstOrDefault(scope => scope.DirectFormatting != null)?.DirectFormatting;
            if (values == null || !values.TryGetValue("version", out string? version) || version != "1") return;
            source.UseDefaultCharacterFormatting = ReadBool(values, "plain") == true;
            source.DirectBold = ReadBool(values, "DirectBold");
            source.DirectItalic = ReadBool(values, "DirectItalic");
            source.DirectUnderlineStyle = ReadEnum<RtfUnderlineStyle>(values, "DirectUnderlineStyle");
            source.FontSize = ReadDouble(values, "FontSize");
            source.FontId = ReadInt(values, "FontId");
            source.ForegroundColorIndex = ReadInt(values, "ForegroundColorIndex");
            source.HighlightColorIndex = ReadInt(values, "HighlightColorIndex");
        }
    }
}

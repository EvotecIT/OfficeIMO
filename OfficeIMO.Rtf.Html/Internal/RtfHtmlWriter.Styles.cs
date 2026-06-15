using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendStylesheetMetadata(StringBuilder builder, RtfDocument document, string newline) {
        if (document.Styles.Count == 0) {
            return;
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int index = 0; index < document.Styles.Count; index++) {
            RtfStyle style = document.Styles[index];
            string prefix = "style." + index.ToString(CultureInfo.InvariantCulture);
            AddStyle(values, prefix, style);
        }

        builder.Append(newline);
        builder.Append("<meta name=\"officeimo-rtf-styles\" content=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append("\">");
    }

    private static void AppendParagraphStyleAttributes(StringBuilder builder, RtfParagraph paragraph) {
        AppendStyleIdAttribute(builder, paragraph.StyleId, "paragraph");
    }

    private static void AppendRunStyleAttributes(StringBuilder builder, RtfRun run) {
        AppendStyleIdAttribute(builder, run.StyleId, "character");
    }

    private static void AppendStyleIdAttribute(StringBuilder builder, int? styleId, string kind) {
        if (!styleId.HasValue) {
            return;
        }

        builder.Append(" data-officeimo-rtf-style-id=\"");
        builder.Append(styleId.Value.ToString(CultureInfo.InvariantCulture));
        builder.Append("\" data-officeimo-rtf-style-kind=\"");
        builder.Append(kind);
        builder.Append('"');
    }

    private static void AddStyle(Dictionary<string, string> values, string prefix, RtfStyle style) {
        AddInt(values, prefix + ".id", style.Id);
        AddString(values, prefix + ".name", style.Name);
        AddEnum(values, prefix + ".kind", (RtfStyleKind?)style.Kind);
        AddNullableInt(values, prefix + ".basedOn", style.BasedOnStyleId);
        AddNullableInt(values, prefix + ".next", style.NextStyleId);
        AddNullableInt(values, prefix + ".linked", style.LinkedStyleId);
        AddKeyCode(values, prefix + ".key", style.KeyCode);
        AddBool(values, prefix + ".additive", style.Additive);
        AddBool(values, prefix + ".autoUpdate", style.AutoUpdate);
        AddBool(values, prefix + ".hidden", style.Hidden);
        AddBool(values, prefix + ".locked", style.Locked);
        AddBool(values, prefix + ".personal", style.Personal);
        AddBool(values, prefix + ".compose", style.Compose);
        AddBool(values, prefix + ".reply", style.Reply);
        AddBool(values, prefix + ".semiHidden", style.SemiHidden);
        AddBool(values, prefix + ".unhideWhenUsed", style.UnhideWhenUsed);
        AddBool(values, prefix + ".quickFormat", style.QuickFormat);
        AddNullableInt(values, prefix + ".priority", style.Priority);
        AddNullableInt(values, prefix + ".revisionSaveId", style.RevisionSaveId);
        AddNullableBool(values, prefix + ".bold", style.Bold);
        AddNullableBool(values, prefix + ".italic", style.Italic);
        AddEnum(values, prefix + ".underlineStyle", style.UnderlineStyle);
        AddNullableDouble(values, prefix + ".fontSize", style.FontSize);
        AddNullableInt(values, prefix + ".fontId", style.FontId);
        AddNullableInt(values, prefix + ".foregroundColor", style.ForegroundColorIndex);
        AddNullableInt(values, prefix + ".highlightColor", style.HighlightColorIndex);
        AddParagraphStyle(values, prefix + ".paragraph", style);
        AddTableStyle(values, prefix + ".table", style);
    }

    private static void AddParagraphStyle(Dictionary<string, string> values, string prefix, RtfStyle style) {
        AddEnum(values, prefix + ".alignment", style.ParagraphAlignment);
        AddEnum(values, prefix + ".direction", style.ParagraphDirection);
        AddNullableInt(values, prefix + ".leftIndent", style.LeftIndentTwips);
        AddNullableInt(values, prefix + ".rightIndent", style.RightIndentTwips);
        AddNullableInt(values, prefix + ".firstLineIndent", style.FirstLineIndentTwips);
        AddNullableInt(values, prefix + ".spaceBefore", style.SpaceBeforeTwips);
        AddNullableInt(values, prefix + ".spaceAfter", style.SpaceAfterTwips);
        AddNullableBool(values, prefix + ".spaceBeforeAuto", style.SpaceBeforeAuto);
        AddNullableBool(values, prefix + ".spaceAfterAuto", style.SpaceAfterAuto);
        AddNullableInt(values, prefix + ".lineSpacing", style.LineSpacingTwips);
        AddNullableBool(values, prefix + ".lineSpacingMultiple", style.LineSpacingMultiple);
        AddNullableInt(values, prefix + ".backgroundColor", style.BackgroundColorIndex);
        AddNullableInt(values, prefix + ".shadingForeground", style.ShadingForegroundColorIndex);
        AddNullableInt(values, prefix + ".shadingPercent", style.ShadingPatternPercent);
        AddEnum(values, prefix + ".shadingPattern", style.ShadingPattern == RtfShadingPattern.None ? (RtfShadingPattern?)null : style.ShadingPattern);
        AddNullableBool(values, prefix + ".pageBreakBefore", style.PageBreakBefore);
        AddNullableBool(values, prefix + ".keepWithNext", style.KeepWithNext);
        AddNullableBool(values, prefix + ".keepLinesTogether", style.KeepLinesTogether);
        AddNullableBool(values, prefix + ".suppressLineNumbers", style.SuppressLineNumbers);
        AddNullableBool(values, prefix + ".autoHyphenation", style.AutoHyphenation);
        AddNullableBool(values, prefix + ".contextualSpacing", style.ContextualSpacing);
        AddNullableBool(values, prefix + ".adjustRightIndent", style.AdjustRightIndent);
        AddNullableBool(values, prefix + ".snapToLineGrid", style.SnapToLineGrid);
        AddNullableBool(values, prefix + ".widowControl", style.WidowControl);
        AddNullableInt(values, prefix + ".outlineLevel", style.OutlineLevel);
        AddBorder(values, prefix + ".border.top", style.TopBorder);
        AddBorder(values, prefix + ".border.left", style.LeftBorder);
        AddBorder(values, prefix + ".border.bottom", style.BottomBorder);
        AddBorder(values, prefix + ".border.right", style.RightBorder);
        AddParagraphFrame(values, prefix + ".frame", style.Frame);
        AddTabStops(values, prefix + ".tab", style.TabStops);
    }

    private static void AddKeyCode(Dictionary<string, string> values, string prefix, RtfStyleKeyCode? keyCode) {
        if (keyCode == null) {
            return;
        }

        AddBool(values, prefix + ".shift", keyCode.Shift);
        AddBool(values, prefix + ".control", keyCode.Control);
        AddBool(values, prefix + ".alt", keyCode.Alt);
        AddNullableInt(values, prefix + ".function", keyCode.FunctionKey);
        AddString(values, prefix + ".key", keyCode.Key);
    }

    private static void AddBorder(Dictionary<string, string> values, string prefix, RtfParagraphBorder border) {
        if (!border.HasAnyValue) {
            return;
        }

        AddEnum(values, prefix + ".style", border.Style == RtfParagraphBorderStyle.None ? (RtfParagraphBorderStyle?)null : border.Style);
        AddNullableInt(values, prefix + ".width", border.Width);
        AddNullableInt(values, prefix + ".color", border.ColorIndex);
    }

    private static void AddTabStops(Dictionary<string, string> values, string prefix, IReadOnlyList<RtfTabStop> tabStops) {
        for (int index = 0; index < tabStops.Count; index++) {
            RtfTabStop tabStop = tabStops[index];
            string tabPrefix = prefix + "." + index.ToString(CultureInfo.InvariantCulture);
            AddInt(values, tabPrefix + ".position", tabStop.PositionTwips);
            AddEnum(values, tabPrefix + ".alignment", (RtfTabAlignment?)tabStop.Alignment);
            AddEnum(values, tabPrefix + ".leader", (RtfTabLeader?)tabStop.Leader);
        }
    }

    private static void AddNullableDouble(Dictionary<string, string> values, string key, double? value) {
        if (value.HasValue) {
            values[key] = value.Value.ToString("R", CultureInfo.InvariantCulture);
        }
    }
}

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendParagraphControlAttributes(StringBuilder builder, RtfParagraph paragraph) {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        AddBool(values, "keepWithNext", paragraph.KeepWithNext);
        AddBool(values, "keepLinesTogether", paragraph.KeepLinesTogether);
        AddBool(values, "suppressLineNumbers", paragraph.SuppressLineNumbers);
        AddNullableBool(values, "autoHyphenation", paragraph.AutoHyphenation);
        AddNullableBool(values, "contextualSpacing", paragraph.ContextualSpacing);
        AddNullableBool(values, "adjustRightIndent", paragraph.AdjustRightIndent);
        AddNullableBool(values, "snapToLineGrid", paragraph.SnapToLineGrid);
        AddNullableBool(values, "widowControl", paragraph.WidowControl);
        AddNullableBool(values, "spaceBeforeAuto", paragraph.SpaceBeforeAuto);
        AddNullableBool(values, "spaceAfterAuto", paragraph.SpaceAfterAuto);
        AddTabStops(values, "tab", paragraph.TabStops);

        if (values.Count == 0) {
            return;
        }

        builder.Append(" data-officeimo-rtf-paragraph-controls=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append('"');
    }
}

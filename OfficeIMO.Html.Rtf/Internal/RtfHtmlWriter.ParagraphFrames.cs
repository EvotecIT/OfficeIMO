namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendParagraphFrameAttributes(StringBuilder builder, RtfParagraph paragraph) {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        AddParagraphFrame(values, "frame", paragraph.Frame);

        if (values.Count == 0) {
            return;
        }

        builder.Append(" data-officeimo-rtf-paragraph-frame=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append('"');
    }

    private static void AddParagraphFrame(Dictionary<string, string> values, string prefix, RtfParagraphFrame frame) {
        if (!frame.HasAnyValue) {
            return;
        }

        AddNullableInt(values, prefix + ".width", frame.WidthTwips);
        AddNullableInt(values, prefix + ".height", frame.HeightTwips);
        AddEnum(values, prefix + ".horizontalAnchor", frame.HorizontalAnchor);
        AddEnum(values, prefix + ".verticalAnchor", frame.VerticalAnchor);
        AddEnum(values, prefix + ".horizontalPosition", frame.HorizontalPosition);
        AddNullableInt(values, prefix + ".horizontalTwips", frame.HorizontalPositionTwips);
        AddEnum(values, prefix + ".verticalPosition", frame.VerticalPosition);
        AddNullableInt(values, prefix + ".verticalTwips", frame.VerticalPositionTwips);
        AddBool(values, prefix + ".anchorLocked", frame.AnchorLocked);
        AddNullableBool(values, prefix + ".noOverlap", frame.NoOverlap);
        AddBool(values, prefix + ".noWrap", frame.NoWrap);
        AddNullableInt(values, prefix + ".wrapDistance", frame.TextWrapDistanceTwips);
        AddNullableInt(values, prefix + ".wrapDistanceHorizontal", frame.TextWrapDistanceHorizontalTwips);
        AddNullableInt(values, prefix + ".wrapDistanceVertical", frame.TextWrapDistanceVerticalTwips);
        AddBool(values, prefix + ".overlayText", frame.OverlayText);
        AddNullableInt(values, prefix + ".dropCapLines", frame.DropCapLines);
        AddEnum(values, prefix + ".dropCapKind", frame.DropCapKind);
    }
}

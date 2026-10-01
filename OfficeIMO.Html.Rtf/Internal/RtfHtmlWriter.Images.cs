namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static bool HasImageCrop(RtfImage image) => image.CropLeftTwips.GetValueOrDefault() != 0 || image.CropTopTwips.GetValueOrDefault() != 0 || image.CropRightTwips.GetValueOrDefault() != 0 || image.CropBottomTwips.GetValueOrDefault() != 0;

    private static void AppendImageMetadata(StringBuilder builder, RtfImage image) {
        var values = new Dictionary<string, string> { ["version"] = "1" };
        AddNullableInt(values, "SourceWidth", image.SourceWidth);
        AddNullableInt(values, "SourceHeight", image.SourceHeight);
        AddNullableInt(values, "DesiredWidthTwips", image.DesiredWidthTwips);
        AddNullableInt(values, "DesiredHeightTwips", image.DesiredHeightTwips);
        AddNullableInt(values, "ScaleXPercent", image.ScaleXPercent);
        AddNullableInt(values, "ScaleYPercent", image.ScaleYPercent);
        AddNullableInt(values, "CropLeftTwips", image.CropLeftTwips);
        AddNullableInt(values, "CropTopTwips", image.CropTopTwips);
        AddNullableInt(values, "CropRightTwips", image.CropRightTwips);
        AddNullableInt(values, "CropBottomTwips", image.CropBottomTwips);
        builder.Append(" data-officeimo-rtf-picture=\"").Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values))).Append('"');
    }
}

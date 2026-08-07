namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendTableRowMetadataAttributes(StringBuilder builder, RtfTableRow row) {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        AddDirectTableRowMetadata(values, "row", row);

        if (values.Count == 0) {
            return;
        }

        builder.Append(" data-officeimo-rtf-row=\"");
        builder.Append(EncodeAttribute(RtfHtmlMetadataCodec.Encode(values)));
        builder.Append('"');
    }

    private static void AddDirectTableRowMetadata(Dictionary<string, string> values, string prefix, RtfTableRow row) {
        AddBool(values, prefix + ".keepTogether", row.KeepTogether);
        AddBool(values, prefix + ".keepWithNext", row.KeepWithNext);
        AddNullableBool(values, prefix + ".autoFit", row.AutoFit);
        AddNullableInt(values, prefix + ".cellGap", row.CellGapTwips);
        AddNullableInt(values, prefix + ".leftIndent", row.LeftIndentTwips);
        AddNullableInt(values, prefix + ".spacing.top", row.SpacingTopTwips);
        AddNullableInt(values, prefix + ".spacing.left", row.SpacingLeftTwips);
        AddNullableInt(values, prefix + ".spacing.bottom", row.SpacingBottomTwips);
        AddNullableInt(values, prefix + ".spacing.right", row.SpacingRightTwips);
        AddBool(values, prefix + ".noOverlap", row.NoOverlap);
        AddEnum(values, prefix + ".horizontalAnchor", row.HorizontalAnchor);
        AddEnum(values, prefix + ".verticalAnchor", row.VerticalAnchor);
        AddEnum(values, prefix + ".horizontalPosition", row.HorizontalPosition);
        AddNullableInt(values, prefix + ".horizontalPositionTwips", row.HorizontalPositionTwips);
        AddEnum(values, prefix + ".verticalPosition", row.VerticalPosition);
        AddNullableInt(values, prefix + ".verticalPositionTwips", row.VerticalPositionTwips);
        AddNullableInt(values, prefix + ".textWrap.left", row.TextWrapLeftTwips);
        AddNullableInt(values, prefix + ".textWrap.right", row.TextWrapRightTwips);
        AddNullableInt(values, prefix + ".textWrap.top", row.TextWrapTopTwips);
        AddNullableInt(values, prefix + ".textWrap.bottom", row.TextWrapBottomTwips);
        AddTableRowBorder(values, prefix + ".border.top", row.TopBorder);
        AddTableRowBorder(values, prefix + ".border.left", row.LeftBorder);
        AddTableRowBorder(values, prefix + ".border.bottom", row.BottomBorder);
        AddTableRowBorder(values, prefix + ".border.right", row.RightBorder);
        AddTableRowBorder(values, prefix + ".border.horizontal", row.HorizontalBorder);
        AddTableRowBorder(values, prefix + ".border.vertical", row.VerticalBorder);
    }

    private static void AddTableRowMetadata(Dictionary<string, string> values, string prefix, RtfTableRow row) {
        AddBool(values, prefix + ".repeatHeader", row.RepeatHeader);
        AddBool(values, prefix + ".keepTogether", row.KeepTogether);
        AddBool(values, prefix + ".keepWithNext", row.KeepWithNext);
        AddNullableBool(values, prefix + ".autoFit", row.AutoFit);
        AddEnum(values, prefix + ".direction", row.Direction);
        AddNullableInt(values, prefix + ".height", row.HeightTwips);
        AddNullableInt(values, prefix + ".cellGap", row.CellGapTwips);
        AddNullableInt(values, prefix + ".leftIndent", row.LeftIndentTwips);
        AddEnum(values, prefix + ".alignment", row.Alignment);
        AddNullableInt(values, prefix + ".preferredWidth", row.PreferredWidth);
        AddEnum(values, prefix + ".preferredWidthUnit", row.PreferredWidthUnit);
        AddNullableInt(values, prefix + ".backgroundColor", row.BackgroundColorIndex);
        AddNullableInt(values, prefix + ".shadingForeground", row.ShadingForegroundColorIndex);
        AddNullableInt(values, prefix + ".shadingPatternValue", row.ShadingPatternValue);
        AddNullableInt(values, prefix + ".shadingPercent", row.ShadingPatternPercent);
        AddEnum(values, prefix + ".shadingPattern", row.ShadingPattern == RtfShadingPattern.None ? (RtfShadingPattern?)null : row.ShadingPattern);
        AddNullableInt(values, prefix + ".padding.top", row.PaddingTopTwips);
        AddNullableInt(values, prefix + ".padding.left", row.PaddingLeftTwips);
        AddNullableInt(values, prefix + ".padding.bottom", row.PaddingBottomTwips);
        AddNullableInt(values, prefix + ".padding.right", row.PaddingRightTwips);
        AddNullableInt(values, prefix + ".spacing.top", row.SpacingTopTwips);
        AddNullableInt(values, prefix + ".spacing.left", row.SpacingLeftTwips);
        AddNullableInt(values, prefix + ".spacing.bottom", row.SpacingBottomTwips);
        AddNullableInt(values, prefix + ".spacing.right", row.SpacingRightTwips);
        AddBool(values, prefix + ".noOverlap", row.NoOverlap);
        AddEnum(values, prefix + ".horizontalAnchor", row.HorizontalAnchor);
        AddEnum(values, prefix + ".verticalAnchor", row.VerticalAnchor);
        AddEnum(values, prefix + ".horizontalPosition", row.HorizontalPosition);
        AddNullableInt(values, prefix + ".horizontalPositionTwips", row.HorizontalPositionTwips);
        AddEnum(values, prefix + ".verticalPosition", row.VerticalPosition);
        AddNullableInt(values, prefix + ".verticalPositionTwips", row.VerticalPositionTwips);
        AddNullableInt(values, prefix + ".textWrap.left", row.TextWrapLeftTwips);
        AddNullableInt(values, prefix + ".textWrap.right", row.TextWrapRightTwips);
        AddNullableInt(values, prefix + ".textWrap.top", row.TextWrapTopTwips);
        AddNullableInt(values, prefix + ".textWrap.bottom", row.TextWrapBottomTwips);
        AddTableRowBorder(values, prefix + ".border.top", row.TopBorder);
        AddTableRowBorder(values, prefix + ".border.left", row.LeftBorder);
        AddTableRowBorder(values, prefix + ".border.bottom", row.BottomBorder);
        AddTableRowBorder(values, prefix + ".border.right", row.RightBorder);
        AddTableRowBorder(values, prefix + ".border.horizontal", row.HorizontalBorder);
        AddTableRowBorder(values, prefix + ".border.vertical", row.VerticalBorder);
    }
}

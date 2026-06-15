namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyTableRowMetadataAttributes(HtmlToken token) {
            Dictionary<string, string> values = RtfHtmlMetadataCodec.Decode(GetAttribute(token, "data-officeimo-rtf-row"));
            if (values.Count == 0 || _row == null) {
                return;
            }

            ApplyDirectTableRowMetadata(values, "row", _row);
        }

        private static void ApplyDirectTableRowMetadata(Dictionary<string, string> values, string prefix, RtfTableRow row) {
            ApplyDirectBool(values, prefix + ".repeatHeader", value => row.RepeatHeader = value);
            ApplyDirectBool(values, prefix + ".keepTogether", value => row.KeepTogether = value);
            ApplyDirectBool(values, prefix + ".keepWithNext", value => row.KeepWithNext = value);
            ApplyDirectNullableBool(values, prefix + ".autoFit", value => row.AutoFit = value);
            ApplyDirectInt(values, prefix + ".cellGap", value => row.CellGapTwips = value);
            ApplyDirectInt(values, prefix + ".leftIndent", value => row.LeftIndentTwips = value);
            ApplyDirectInt(values, prefix + ".spacing.top", value => row.SpacingTopTwips = value);
            ApplyDirectInt(values, prefix + ".spacing.left", value => row.SpacingLeftTwips = value);
            ApplyDirectInt(values, prefix + ".spacing.bottom", value => row.SpacingBottomTwips = value);
            ApplyDirectInt(values, prefix + ".spacing.right", value => row.SpacingRightTwips = value);
            ApplyDirectBool(values, prefix + ".noOverlap", value => row.NoOverlap = value);
            ApplyDirectEnum<RtfTableHorizontalAnchor>(values, prefix + ".horizontalAnchor", value => row.HorizontalAnchor = value);
            ApplyDirectEnum<RtfTableVerticalAnchor>(values, prefix + ".verticalAnchor", value => row.VerticalAnchor = value);
            ApplyDirectEnum<RtfTableHorizontalPosition>(values, prefix + ".horizontalPosition", value => row.HorizontalPosition = value);
            ApplyDirectInt(values, prefix + ".horizontalPositionTwips", value => row.HorizontalPositionTwips = value);
            ApplyDirectEnum<RtfTableVerticalPosition>(values, prefix + ".verticalPosition", value => row.VerticalPosition = value);
            ApplyDirectInt(values, prefix + ".verticalPositionTwips", value => row.VerticalPositionTwips = value);
            ApplyDirectInt(values, prefix + ".textWrap.left", value => row.TextWrapLeftTwips = value);
            ApplyDirectInt(values, prefix + ".textWrap.right", value => row.TextWrapRightTwips = value);
            ApplyDirectInt(values, prefix + ".textWrap.top", value => row.TextWrapTopTwips = value);
            ApplyDirectInt(values, prefix + ".textWrap.bottom", value => row.TextWrapBottomTwips = value);
            ApplyDirectTableRowBorder(values, prefix + ".border.top", row.TopBorder);
            ApplyDirectTableRowBorder(values, prefix + ".border.left", row.LeftBorder);
            ApplyDirectTableRowBorder(values, prefix + ".border.bottom", row.BottomBorder);
            ApplyDirectTableRowBorder(values, prefix + ".border.right", row.RightBorder);
            ApplyDirectTableRowBorder(values, prefix + ".border.horizontal", row.HorizontalBorder);
            ApplyDirectTableRowBorder(values, prefix + ".border.vertical", row.VerticalBorder);
        }

        private static void ApplyTableRowMetadata(Dictionary<string, string> values, string prefix, RtfTableRow row) {
            row.RepeatHeader = ReadBool(values, prefix + ".repeatHeader") == true;
            row.KeepTogether = ReadBool(values, prefix + ".keepTogether") == true;
            row.KeepWithNext = ReadBool(values, prefix + ".keepWithNext") == true;
            row.AutoFit = ReadBool(values, prefix + ".autoFit");
            row.Direction = ReadEnum<RtfTableRowDirection>(values, prefix + ".direction");
            row.HeightTwips = ReadInt(values, prefix + ".height");
            row.CellGapTwips = ReadInt(values, prefix + ".cellGap");
            row.LeftIndentTwips = ReadInt(values, prefix + ".leftIndent");
            row.Alignment = ReadEnum<RtfTableAlignment>(values, prefix + ".alignment");
            row.PreferredWidth = ReadInt(values, prefix + ".preferredWidth");
            row.PreferredWidthUnit = ReadEnum<RtfTableWidthUnit>(values, prefix + ".preferredWidthUnit");
            row.BackgroundColorIndex = ReadInt(values, prefix + ".backgroundColor");
            row.ShadingForegroundColorIndex = ReadInt(values, prefix + ".shadingForeground");
            row.ShadingPatternValue = ReadInt(values, prefix + ".shadingPatternValue");
            row.ShadingPatternPercent = ReadInt(values, prefix + ".shadingPercent");
            row.ShadingPattern = ReadEnum(values, prefix + ".shadingPattern", RtfShadingPattern.None);
            row.PaddingTopTwips = ReadInt(values, prefix + ".padding.top");
            row.PaddingLeftTwips = ReadInt(values, prefix + ".padding.left");
            row.PaddingBottomTwips = ReadInt(values, prefix + ".padding.bottom");
            row.PaddingRightTwips = ReadInt(values, prefix + ".padding.right");
            row.SpacingTopTwips = ReadInt(values, prefix + ".spacing.top");
            row.SpacingLeftTwips = ReadInt(values, prefix + ".spacing.left");
            row.SpacingBottomTwips = ReadInt(values, prefix + ".spacing.bottom");
            row.SpacingRightTwips = ReadInt(values, prefix + ".spacing.right");
            row.NoOverlap = ReadBool(values, prefix + ".noOverlap") == true;
            row.HorizontalAnchor = ReadEnum<RtfTableHorizontalAnchor>(values, prefix + ".horizontalAnchor");
            row.VerticalAnchor = ReadEnum<RtfTableVerticalAnchor>(values, prefix + ".verticalAnchor");
            row.HorizontalPosition = ReadEnum<RtfTableHorizontalPosition>(values, prefix + ".horizontalPosition");
            row.HorizontalPositionTwips = ReadInt(values, prefix + ".horizontalPositionTwips");
            row.VerticalPosition = ReadEnum<RtfTableVerticalPosition>(values, prefix + ".verticalPosition");
            row.VerticalPositionTwips = ReadInt(values, prefix + ".verticalPositionTwips");
            row.TextWrapLeftTwips = ReadInt(values, prefix + ".textWrap.left");
            row.TextWrapRightTwips = ReadInt(values, prefix + ".textWrap.right");
            row.TextWrapTopTwips = ReadInt(values, prefix + ".textWrap.top");
            row.TextWrapBottomTwips = ReadInt(values, prefix + ".textWrap.bottom");
            ApplyTableRowBorder(values, prefix + ".border.top", row.TopBorder);
            ApplyTableRowBorder(values, prefix + ".border.left", row.LeftBorder);
            ApplyTableRowBorder(values, prefix + ".border.bottom", row.BottomBorder);
            ApplyTableRowBorder(values, prefix + ".border.right", row.RightBorder);
            ApplyTableRowBorder(values, prefix + ".border.horizontal", row.HorizontalBorder);
            ApplyTableRowBorder(values, prefix + ".border.vertical", row.VerticalBorder);
        }

        private static void ApplyDirectTableRowBorder(Dictionary<string, string> values, string prefix, RtfTableRowBorder border) {
            if (!values.Keys.Any(key => key.StartsWith(prefix + ".", StringComparison.Ordinal))) {
                return;
            }

            ApplyTableRowBorder(values, prefix, border);
        }

        private static void ApplyDirectBool(Dictionary<string, string> values, string key, Action<bool> assign) {
            if (values.ContainsKey(key)) {
                assign(ReadBool(values, key) == true);
            }
        }

        private static void ApplyDirectNullableBool(Dictionary<string, string> values, string key, Action<bool?> assign) {
            if (values.ContainsKey(key)) {
                assign(ReadBool(values, key));
            }
        }

        private static void ApplyDirectInt(Dictionary<string, string> values, string key, Action<int?> assign) {
            if (values.ContainsKey(key)) {
                assign(ReadInt(values, key));
            }
        }

        private static void ApplyDirectEnum<T>(Dictionary<string, string> values, string key, Action<T?> assign) where T : struct {
            if (values.ContainsKey(key)) {
                assign(ReadEnum<T>(values, key));
            }
        }
    }
}

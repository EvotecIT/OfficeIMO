using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private static void ApplyTableStyle(Dictionary<string, string> values, string prefix, RtfStyle style) {
            if (style.Kind != RtfStyleKind.Table) {
                return;
            }

            RtfTableRow row = style.TableRowFormat;
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
            ApplyTableStyleCells(values, prefix + ".cell", row);
        }

        private static void ApplyTableStyleCells(Dictionary<string, string> values, string prefix, RtfTableRow row) {
            for (int index = 0; ; index++) {
                string cellPrefix = prefix + "." + index.ToString(CultureInfo.InvariantCulture);
                if (!HasTableStyleCell(values, cellPrefix)) {
                    break;
                }

                RtfTableCell cell = row.AddCell(ReadInt(values, cellPrefix + ".rightBoundary"));
                cell.HorizontalMerge = ReadEnum(values, cellPrefix + ".horizontalMerge", RtfTableCellMerge.None);
                cell.VerticalMerge = ReadEnum(values, cellPrefix + ".verticalMerge", RtfTableCellMerge.None);
                cell.BackgroundColorIndex = ReadInt(values, cellPrefix + ".backgroundColor");
                cell.ShadingForegroundColorIndex = ReadInt(values, cellPrefix + ".shadingForeground");
                cell.ShadingPatternPercent = ReadInt(values, cellPrefix + ".shadingPercent");
                cell.ShadingPattern = ReadEnum(values, cellPrefix + ".shadingPattern", RtfShadingPattern.None);
                cell.VerticalAlignment = ReadEnum<RtfTableCellVerticalAlignment>(values, cellPrefix + ".verticalAlignment");
                cell.TextFlow = ReadEnum<RtfTableCellTextFlow>(values, cellPrefix + ".textFlow");
                cell.PreferredWidth = ReadInt(values, cellPrefix + ".preferredWidth");
                cell.PreferredWidthUnit = ReadEnum<RtfTableWidthUnit>(values, cellPrefix + ".preferredWidthUnit");
                cell.HideCellMark = ReadBool(values, cellPrefix + ".hideCellMark") == true;
                cell.NoWrap = ReadBool(values, cellPrefix + ".noWrap") == true;
                cell.FitText = ReadBool(values, cellPrefix + ".fitText") == true;
                cell.PaddingTopTwips = ReadInt(values, cellPrefix + ".padding.top");
                cell.PaddingLeftTwips = ReadInt(values, cellPrefix + ".padding.left");
                cell.PaddingBottomTwips = ReadInt(values, cellPrefix + ".padding.bottom");
                cell.PaddingRightTwips = ReadInt(values, cellPrefix + ".padding.right");
                ApplyTableCellBorder(values, cellPrefix + ".border.top", cell.TopBorder);
                ApplyTableCellBorder(values, cellPrefix + ".border.left", cell.LeftBorder);
                ApplyTableCellBorder(values, cellPrefix + ".border.bottom", cell.BottomBorder);
                ApplyTableCellBorder(values, cellPrefix + ".border.right", cell.RightBorder);
                ApplyTableCellBorder(values, cellPrefix + ".border.diagonalDown", cell.TopLeftToBottomRightBorder);
                ApplyTableCellBorder(values, cellPrefix + ".border.diagonalUp", cell.TopRightToBottomLeftBorder);
            }
        }

        private static bool HasTableStyleCell(Dictionary<string, string> values, string prefix) {
            return values.Keys.Any(key => key.StartsWith(prefix + ".", StringComparison.Ordinal));
        }

        private static void ApplyTableRowBorder(Dictionary<string, string> values, string prefix, RtfTableRowBorder border) {
            border.Style = ReadEnum(values, prefix + ".style", RtfTableCellBorderStyle.None);
            border.Width = ReadInt(values, prefix + ".width");
            border.ColorIndex = ReadInt(values, prefix + ".color");
        }

        private static void ApplyTableCellBorder(Dictionary<string, string> values, string prefix, RtfTableCellBorder border) {
            border.Style = ReadEnum(values, prefix + ".style", RtfTableCellBorderStyle.None);
            border.Width = ReadInt(values, prefix + ".width");
            border.ColorIndex = ReadInt(values, prefix + ".color");
        }
    }
}

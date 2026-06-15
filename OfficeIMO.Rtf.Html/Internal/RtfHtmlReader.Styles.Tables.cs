using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private static void ApplyTableStyle(Dictionary<string, string> values, string prefix, RtfStyle style) {
            if (style.Kind != RtfStyleKind.Table) {
                return;
            }

            RtfTableRow row = style.TableRowFormat;
            ApplyTableRowMetadata(values, prefix, row);
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

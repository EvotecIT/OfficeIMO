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
                if (!HasTableCellMetadata(values, cellPrefix)) {
                    break;
                }

                RtfTableCell cell = row.AddCell(ReadInt(values, cellPrefix + ".rightBoundary"));
                ApplyTableCellMetadata(values, cellPrefix, cell);
            }
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

using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlWriter {
    private static void AddTableStyle(Dictionary<string, string> values, string prefix, RtfStyle style) {
        if (style.Kind != RtfStyleKind.Table) {
            return;
        }

        RtfTableRow row = style.TableRowFormat;
        AddTableRowMetadata(values, prefix, row);
        AddTableStyleCells(values, prefix + ".cell", row.Cells);
    }

    private static void AddTableStyleCells(Dictionary<string, string> values, string prefix, IReadOnlyList<RtfTableCell> cells) {
        for (int index = 0; index < cells.Count; index++) {
            RtfTableCell cell = cells[index];
            string cellPrefix = prefix + "." + index.ToString(CultureInfo.InvariantCulture);
            AddTableCellMetadata(values, cellPrefix, cell);
        }
    }

    private static void AddTableRowBorder(Dictionary<string, string> values, string prefix, RtfTableRowBorder border) {
        if (!border.HasAnyValue) {
            return;
        }

        AddEnum(values, prefix + ".style", border.Style == RtfTableCellBorderStyle.None ? (RtfTableCellBorderStyle?)null : border.Style);
        AddNullableInt(values, prefix + ".width", border.Width);
        AddNullableInt(values, prefix + ".color", border.ColorIndex);
    }

    private static void AddTableCellBorder(Dictionary<string, string> values, string prefix, RtfTableCellBorder border) {
        if (!border.HasAnyValue) {
            return;
        }

        AddEnum(values, prefix + ".style", border.Style == RtfTableCellBorderStyle.None ? (RtfTableCellBorderStyle?)null : border.Style);
        AddNullableInt(values, prefix + ".width", border.Width);
        AddNullableInt(values, prefix + ".color", border.ColorIndex);
    }
}

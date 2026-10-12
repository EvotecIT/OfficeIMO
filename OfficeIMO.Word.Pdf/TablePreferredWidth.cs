using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    /// <summary>Reads preferred table measurements without assuming the stored value is an integer.</summary>
    internal static class TablePreferredWidth {
        internal static TableWidthUnitValues ResolveType(TableWidthType? width) {
            string? rawWidth = width?.Width?.Value;
            string value = rawWidth?.Trim() ?? string.Empty;
            bool percentage = value.EndsWith("%", StringComparison.Ordinal);
            TableWidthUnitValues type = width?.Type?.Value ?? TableWidthUnitValues.Dxa;
            if (value.Length == 0) return TableWidthUnitValues.Auto;
            string numeric = percentage ? value.Substring(0, value.Length - 1) : value;
            if (double.TryParse(numeric, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) && number == 0D)
                return TableWidthUnitValues.Auto;
            if (percentage) return TableWidthUnitValues.Pct;
            return type == TableWidthUnitValues.Nil ? TableWidthUnitValues.Auto : type;
        }

        /// <summary>Cell measurements retain zero absolute preferences; an absent measurement is automatic.</summary>
        internal static TableWidthUnitValues ResolveCellType(TableCellWidth? width) {
            if (width == null) return TableWidthUnitValues.Auto;
            string value = width.Width?.Value?.Trim() ?? string.Empty;
            if (value.Length == 0) return TableWidthUnitValues.Dxa;
            if (value.EndsWith("%", StringComparison.Ordinal)) return TableWidthUnitValues.Pct;
            return width.Type?.Value ?? TableWidthUnitValues.Dxa;
        }

        internal static double? AbsoluteCellPoints(WordTableCell cell) {
            TableCellWidth? width = cell._tableCellProperties?.TableCellWidth;
            if (ResolveCellType(width) != TableWidthUnitValues.Dxa) return null;
            string? value = width?.Width?.Value;
            if (string.IsNullOrWhiteSpace(value)) return 0D;
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips) && twips >= 0
                ? twips / 20D : null;
        }

        internal static bool PreservesCellGrid(WordTableCell cell) =>
            ResolveCellType(cell._tableCellProperties?.TableCellWidth) == TableWidthUnitValues.Nil;

        internal static bool HasAbsoluteCellPreference(WordTableCell cell) =>
            AbsoluteCellPoints(cell).HasValue || PreservesCellGrid(cell);

        internal static bool CanDisableCellWrapping(WordTableCell cell) {
            TableWidthUnitValues type = ResolveCellType(cell._tableCellProperties?.TableCellWidth);
            return type == TableWidthUnitValues.Auto || type == TableWidthUnitValues.Pct;
        }

        internal static double? Percentage(string? rawWidth) {
            if (string.IsNullOrWhiteSpace(rawWidth)) return null;
            string value = rawWidth!.Trim();
            bool percentage = value.EndsWith("%", StringComparison.Ordinal);
            string numeric = percentage ? value.Substring(0, value.Length - 1) : value;
            if (!percentage) {
                return int.TryParse(numeric, NumberStyles.Integer, CultureInfo.InvariantCulture, out int integer) && integer > 0
                    ? integer / 5000D : null;
            }
            if (!double.TryParse(numeric, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) ||
                number <= 0D || double.IsNaN(number) || double.IsInfinity(number)) return null;
            return number / 100D;
        }
    }
}

using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>Reads preferred table measurements without assuming the stored value is an integer.</summary>
    internal static class WordTableWidthNative {
        // Read lexical type tokens without EnumValue.Value, which throws on unknown values.
        internal static TableWidthUnitValues? NativeType(TableWidthType? width) {
            string? token = width?.Type?.InnerText;
            if (token == null || token == "dxa") return TableWidthUnitValues.Dxa;
            if (token == "pct") return TableWidthUnitValues.Pct;
            if (token == "auto") return TableWidthUnitValues.Auto;
            if (token == "nil") return TableWidthUnitValues.Nil;
            return null;
        }

        /// <summary>Reads authored preferences without changing lexical or partial native declarations.
        /// Missing measurements remain in XML and read as zero; layout uses its separate table/cell rules.
        /// Unsupported declarations throw only when requesting the public value, never on native preservation.</summary>
        internal static WordTableWidth? Read(TableWidthType? width) {
            if (width == null) return null;
            TableWidthUnitValues? type = NativeType(width);
            string value = width.Width?.Value?.Trim() ?? string.Empty;
            if (!type.HasValue) throw new NotSupportedException("Unsupported Word preferred-width type: " + width.Type?.InnerText);
            if (value.EndsWith("%", StringComparison.Ordinal)) {
                string numeric = value.Substring(0, value.Length - 1);
                if (double.TryParse(numeric, NumberStyles.Float, CultureInfo.InvariantCulture, out double percent) &&
                    percent >= 0 && !double.IsNaN(percent) && !double.IsInfinity(percent) && percent <= int.MaxValue / 50d)
                    return WordTableWidth.Percent(percent);
                throw new NotSupportedException("Unsupported Word percentage width: " + value);
            }
            if (type == TableWidthUnitValues.Auto) return WordTableWidth.Auto;
            if (type == TableWidthUnitValues.Nil) return WordTableWidth.Nil;
            int number = 0;
            if (value.Length != 0 && (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out number) || number < 0))
                throw new NotSupportedException("Unsupported Word preferred width: " + value);
            return type == TableWidthUnitValues.Pct ? WordTableWidth.Percent(number / 50d)
                : WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)number * 635)));
        }

        /// <summary>Writes both native attributes before any dependent layout refresh.</summary>
        internal static void Write(TableWidthType native, WordTableWidth width) {
            native.Type = width.Kind == WordTableWidthKind.Length ? TableWidthUnitValues.Dxa
                : width.Kind == WordTableWidthKind.Percent ? TableWidthUnitValues.Pct
                : width.Kind == WordTableWidthKind.Auto ? TableWidthUnitValues.Auto : TableWidthUnitValues.Nil;
            native.Width = width.ToNativeValue().ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>Reads a numeric native measurement without resolving its unit or inventing a missing value.</summary>
        internal static int? NumericValue(TableWidthType? width) =>
            int.TryParse(width?.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : (int?)null;

        internal static TableWidthUnitValues? ResolveType(TableWidthType? width) {
            string? rawWidth = width?.Width?.Value;
            string value = rawWidth?.Trim() ?? string.Empty;
            bool percentage = value.EndsWith("%", StringComparison.Ordinal);
            TableWidthUnitValues? type = NativeType(width);
            if (!type.HasValue) return null;
            if (value.Length == 0) return TableWidthUnitValues.Auto;
            string numeric = percentage ? value.Substring(0, value.Length - 1) : value;
            if (double.TryParse(numeric, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) && number == 0D)
                return TableWidthUnitValues.Auto;
            if (percentage) return TableWidthUnitValues.Pct;
            return type == TableWidthUnitValues.Nil ? TableWidthUnitValues.Auto : type;
        }

        /// <summary>Cell measurements retain zero absolute preferences; an absent measurement is automatic.</summary>
        internal static TableWidthUnitValues? ResolveCellType(TableCellWidth? width) {
            if (width == null) return TableWidthUnitValues.Auto;
            if (!NativeType(width).HasValue) return null;
            string value = width.Width?.Value?.Trim() ?? string.Empty;
            if (value.Length == 0) return TableWidthUnitValues.Dxa;
            if (value.EndsWith("%", StringComparison.Ordinal)) return TableWidthUnitValues.Pct;
            return NativeType(width);
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
            TableWidthUnitValues? type = ResolveCellType(cell._tableCellProperties?.TableCellWidth);
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

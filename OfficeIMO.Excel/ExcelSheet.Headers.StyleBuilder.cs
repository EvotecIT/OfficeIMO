using System.Globalization;

namespace OfficeIMO.Excel
{
    /// <summary>
    /// Builder for applying styles/number formats to a column resolved by header.
    /// </summary>
    public sealed class ColumnStyleByHeaderBuilder
    {
        private readonly ExcelSheet _sheet;
        private readonly int _colIndex;
        private readonly int _startRow;
        private readonly int _endRow;

        internal ColumnStyleByHeaderBuilder(ExcelSheet sheet, int colIndex, int startRow, int endRow)
        {
            _sheet = sheet;
            _colIndex = colIndex;
            _startRow = startRow;
            _endRow = endRow;
        }

        private ColumnStyleByHeaderBuilder ApplyFormat(string numberFormat)
        {
            for (int r = _startRow; r <= _endRow; r++)
                _sheet.FormatCell(r, _colIndex, numberFormat);
            return this;
        }

        /// <summary>
        /// Applies a numeric format with optional decimal places.
        /// </summary>
        /// <param name="decimals">Number of decimal places.</param>
        public ColumnStyleByHeaderBuilder Number(int decimals = 0)
            => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Decimal, decimals));

        /// <summary>
        /// Applies an integer number format.
        /// </summary>
        public ColumnStyleByHeaderBuilder Integer() => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Integer));

        /// <summary>
        /// Applies a percentage format.
        /// </summary>
        /// <param name="decimals">Number of decimal places.</param>
        public ColumnStyleByHeaderBuilder Percent(int decimals = 0)
            => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Percent, decimals));

        /// <summary>
        /// Applies a currency format using the specified culture.
        /// </summary>
        /// <param name="decimals">Number of decimal places.</param>
        /// <param name="culture">Culture for currency symbol.</param>
        public ColumnStyleByHeaderBuilder Currency(int decimals = 2, CultureInfo? culture = null)
            => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Currency, decimals, culture));

        /// <summary>
        /// Applies a date format using the provided pattern.
        /// </summary>
        /// <param name="pattern">Date format pattern.</param>
        public ColumnStyleByHeaderBuilder Date(string pattern = "yyyy-mm-dd") => ApplyFormat(pattern);

        /// <summary>
        /// Applies a date and time format using the provided pattern.
        /// </summary>
        /// <param name="pattern">DateTime format pattern.</param>
        public ColumnStyleByHeaderBuilder DateTime(string pattern = "yyyy-mm-dd hh:mm:ss") => ApplyFormat(pattern);

        /// <summary>
        /// Applies a time format.
        /// </summary>
        public ColumnStyleByHeaderBuilder Time() => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Time));

        /// <summary>
        /// Applies a duration format in hours.
        /// </summary>
        public ColumnStyleByHeaderBuilder DurationHours() => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.DurationHours));

        /// <summary>
        /// Applies a text format.
        /// </summary>
        public ColumnStyleByHeaderBuilder Text() => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Text));

        /// <summary>
        /// Applies a custom number format string.
        /// </summary>
        /// <param name="format">Number format pattern.</param>
        public ColumnStyleByHeaderBuilder NumberFormat(string format) => ApplyFormat(format);

        /// <summary>
        /// Makes all cells in the column bold.
        /// </summary>
        public ColumnStyleByHeaderBuilder Bold()
        {
            for (int r = _startRow; r <= _endRow; r++)
                _sheet.CellBold(r, _colIndex, bold: true);
            return this;
        }

        /// <summary>
        /// Applies a solid background fill to the column. Accepts #RRGGBB or #AARRGGBB.
        /// </summary>
        public ColumnStyleByHeaderBuilder Background(string hexColor)
        {
            for (int r = _startRow; r <= _endRow; r++)
                _sheet.CellBackground(r, _colIndex, hexColor);
            return this;
        }

        /// <summary>
        /// Applies a solid background fill to the column using an ImageSharp color.
        /// </summary>
        /// <param name="color">Fill color.</param>
        public ColumnStyleByHeaderBuilder Background(SixLabors.ImageSharp.Color color)
        {
            for (int r = _startRow; r <= _endRow; r++)
                _sheet.CellBackground(r, _colIndex, color);
            return this;
        }

        /// <summary>
        /// Applies a background fill to cells in this column when their text equals the specified value.
        /// Comparison is case-insensitive by default.
        /// </summary>
        public ColumnStyleByHeaderBuilder BackgroundWhenTextEquals(string text, string hexColor, bool caseInsensitive = true)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(hexColor)) return this;
            var comparison = caseInsensitive ? System.StringComparison.OrdinalIgnoreCase : System.StringComparison.Ordinal;
            for (int r = _startRow; r <= _endRow; r++)
            {
                if (_sheet.TryGetCellText(r, _colIndex, out var value))
                {
                    if (string.Equals(value, text, comparison))
                        _sheet.CellBackground(r, _colIndex, hexColor);
                }
            }
            return this;
        }

        /// <summary>
        /// Applies background fills to cells based on a text→color mapping.
        /// Keys are compared case-insensitively by default.
        /// </summary>
        public ColumnStyleByHeaderBuilder BackgroundByTextMap(System.Collections.Generic.IDictionary<string, string> map, bool caseInsensitive = true)
        {
            if (map == null || map.Count == 0) return this;
            var dict = caseInsensitive ? new System.Collections.Generic.Dictionary<string, string>(map, System.StringComparer.OrdinalIgnoreCase)
                                       : new System.Collections.Generic.Dictionary<string, string>(map);
            for (int r = _startRow; r <= _endRow; r++)
            {
                if (_sheet.TryGetCellText(r, _colIndex, out var value) && !string.IsNullOrEmpty(value) && dict.TryGetValue(value, out var color))
                {
                    _sheet.CellBackground(r, _colIndex, color);
                }
            }
            return this;
        }

        /// <summary>
        /// Overload that accepts SixLabors colors for convenience.
        /// </summary>
        public ColumnStyleByHeaderBuilder BackgroundByTextMap(System.Collections.Generic.IDictionary<string, SixLabors.ImageSharp.Color> map, bool caseInsensitive = true)
        {
            if (map == null || map.Count == 0) return this;
            var hex = new System.Collections.Generic.Dictionary<string, string>(caseInsensitive ? System.StringComparer.OrdinalIgnoreCase : System.StringComparer.Ordinal);
            foreach (var kv in map)
                hex[kv.Key] = OfficeIMO.Excel.ExcelColor.ToArgbHex(kv.Value);
            return BackgroundByTextMap(hex, caseInsensitive);
        }

        /// <summary>
        /// Makes the cell bold when its text equals the specified value.
        /// </summary>
        public ColumnStyleByHeaderBuilder BoldWhenTextEquals(string text, bool caseInsensitive = true)
        {
            if (string.IsNullOrEmpty(text)) return this;
            var comparison = caseInsensitive ? System.StringComparison.OrdinalIgnoreCase : System.StringComparison.Ordinal;
            for (int r = _startRow; r <= _endRow; r++)
            {
                if (_sheet.TryGetCellText(r, _colIndex, out var value) && string.Equals(value, text, comparison))
                    _sheet.CellBold(r, _colIndex, true);
            }
            return this;
        }

        /// <summary>
        /// Makes the cell bold when its text is in the provided set.
        /// </summary>
        public ColumnStyleByHeaderBuilder BoldByTextSet(System.Collections.Generic.ISet<string> values, bool caseInsensitive = true)
        {
            if (values == null || values.Count == 0) return this;
            var set = caseInsensitive ? new System.Collections.Generic.HashSet<string>(values, System.StringComparer.OrdinalIgnoreCase)
                                      : new System.Collections.Generic.HashSet<string>(values);
            for (int r = _startRow; r <= _endRow; r++)
            {
                if (_sheet.TryGetCellText(r, _colIndex, out var value) && !string.IsNullOrEmpty(value) && set.Contains(value))
                    _sheet.CellBold(r, _colIndex, true);
            }
            return this;
        }
    }
}

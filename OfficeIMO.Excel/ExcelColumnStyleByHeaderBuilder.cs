using System.Globalization;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Builder for applying styles/number formats to a column resolved by header.
    /// </summary>
    public sealed class ExcelColumnStyleByHeaderBuilder {
        private readonly ExcelSheet _sheet;
        private readonly int _colIndex;
        private readonly int _startRow;
        private readonly int _endRow;

        internal ExcelColumnStyleByHeaderBuilder(ExcelSheet sheet, int colIndex, int startRow, int endRow) {
            _sheet = sheet;
            _colIndex = colIndex;
            _startRow = startRow;
            _endRow = endRow;
        }

        private ExcelColumnStyleByHeaderBuilder ApplyFormat(string numberFormat) {
            if (_colIndex <= 0 || _startRow > _endRow) {
                return this;
            }

            if (_startRow > 1
                && _sheet.Document.TrySetDirectTabularSaveCandidateColumnNumberFormat(_sheet, _colIndex, numberFormat)) {
                return this;
            }

            string column = ExcelA1.ColumnIndexToLetters(_colIndex);
            _sheet.FormatRange(
                column + _startRow.ToString(CultureInfo.InvariantCulture) + ":" +
                column + _endRow.ToString(CultureInfo.InvariantCulture),
                numberFormat);
            return this;
        }

        /// <summary>
        /// Applies a numeric format with optional decimal places.
        /// </summary>
        /// <param name="decimals">Number of decimal places.</param>
        public ExcelColumnStyleByHeaderBuilder Number(int decimals = 0)
            => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Decimal, decimals));

        /// <summary>
        /// Applies an integer number format.
        /// </summary>
        public ExcelColumnStyleByHeaderBuilder Integer() => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Integer));

        /// <summary>
        /// Applies a percentage format.
        /// </summary>
        /// <param name="decimals">Number of decimal places.</param>
        public ExcelColumnStyleByHeaderBuilder Percent(int decimals = 0)
            => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Percent, decimals));

        /// <summary>
        /// Applies a currency format using the specified culture.
        /// </summary>
        /// <param name="decimals">Number of decimal places.</param>
        /// <param name="culture">Culture for currency symbol.</param>
        public ExcelColumnStyleByHeaderBuilder Currency(int decimals = 2, CultureInfo? culture = null)
            => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Currency, decimals, culture));

        /// <summary>
        /// Applies a date format using the provided pattern.
        /// </summary>
        /// <param name="pattern">Date format pattern.</param>
        public ExcelColumnStyleByHeaderBuilder Date(string pattern = "yyyy-mm-dd") => ApplyFormat(pattern);

        /// <summary>
        /// Applies a date and time format using the provided pattern.
        /// </summary>
        /// <param name="pattern">DateTime format pattern.</param>
        public ExcelColumnStyleByHeaderBuilder DateTime(string pattern = "yyyy-mm-dd hh:mm:ss") => ApplyFormat(pattern);

        /// <summary>
        /// Applies a time format.
        /// </summary>
        public ExcelColumnStyleByHeaderBuilder Time() => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Time));

        /// <summary>
        /// Applies a duration format in hours.
        /// </summary>
        public ExcelColumnStyleByHeaderBuilder DurationHours() => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.DurationHours));

        /// <summary>
        /// Applies a text format.
        /// </summary>
        public ExcelColumnStyleByHeaderBuilder Text() => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Text));

        /// <summary>Applies a custom number format without materializing a pending direct tabular export.</summary>
        public ExcelColumnStyleByHeaderBuilder SetNumberFormat(string format) => ApplyFormat(format);

        /// <summary>Sets or clears bold text.</summary>
        public ExcelColumnStyleByHeaderBuilder SetBold(bool bold = true) =>
            Apply((row, column) => _sheet.CellBold(row, column, bold));

        /// <summary>Sets or clears italic text.</summary>
        public ExcelColumnStyleByHeaderBuilder SetItalic(bool italic = true) =>
            Apply((row, column) => _sheet.CellItalic(row, column, italic));

        /// <summary>Applies single underline.</summary>
        public ExcelColumnStyleByHeaderBuilder SetUnderline() => SetUnderline(true);

        /// <summary>Sets or clears single underline.</summary>
        public ExcelColumnStyleByHeaderBuilder SetUnderline(bool underline) =>
            Apply((row, column) => _sheet.CellUnderline(row, column, underline));

        /// <summary>Sets the native Excel underline style.</summary>
        public ExcelColumnStyleByHeaderBuilder SetUnderline(ExcelUnderlineStyle underlineStyle) =>
            Apply((row, column) => _sheet.CellUnderline(row, column, underlineStyle));

        /// <summary>Sets or clears strikethrough.</summary>
        public ExcelColumnStyleByHeaderBuilder SetStrikethrough(bool strikethrough = true) =>
            Apply((row, column) => _sheet.CellStrikethrough(row, column, strikethrough));

        /// <summary>Sets the typeface family name.</summary>
        public ExcelColumnStyleByHeaderBuilder SetFontFamily(string fontFamily) =>
            Apply((row, column) => _sheet.CellFontName(row, column, fontFamily));

        /// <summary>Sets a positive physical font size.</summary>
        public ExcelColumnStyleByHeaderBuilder SetFontSize(OfficeLength fontSize) =>
            Apply((row, column) => _sheet.CellFontSize(row, column, fontSize));

        /// <summary>Sets the literal or theme-based text color.</summary>
        public ExcelColumnStyleByHeaderBuilder SetColor(OfficeColor color) =>
            Apply((row, column) => _sheet.CellFontColor(row, column, color));

        /// <summary>Sets the literal or theme-based solid fill color.</summary>
        public ExcelColumnStyleByHeaderBuilder SetFillColor(OfficeColor color) =>
            Apply((row, column) => _sheet.CellBackground(row, column, color));

        /// <summary>Sets horizontal alignment.</summary>
        public ExcelColumnStyleByHeaderBuilder SetHorizontalAlignment(ExcelHorizontalAlignment alignment) =>
            Apply((row, column) => _sheet.CellAlign(row, column, alignment));

        /// <summary>Sets vertical alignment.</summary>
        public ExcelColumnStyleByHeaderBuilder SetVerticalAlignment(ExcelVerticalAlignment alignment) =>
            Apply((row, column) => _sheet.CellVerticalAlign(row, column, alignment));

        /// <summary>Sets baseline, superscript or subscript text alignment.</summary>
        public ExcelColumnStyleByHeaderBuilder SetVerticalTextAlignment(ExcelVerticalTextAlignment alignment) =>
            Apply((row, column) => _sheet.CellVerticalTextAlignment(row, column, alignment));

        /// <summary>Sets or clears text wrapping.</summary>
        public ExcelColumnStyleByHeaderBuilder SetWrapText(bool wrapText = true) =>
            Apply((row, column) => _sheet.CellWrapText(row, column, wrapText));

        /// <summary>Sets text colors by matching cell text to a color map.</summary>
        public ExcelColumnStyleByHeaderBuilder SetColorByTextMap(IDictionary<string, OfficeColor> map, bool caseInsensitive = true) =>
            ApplyColorMap(map, caseInsensitive, (row, column, color) => _sheet.CellFontColor(row, column, color));

        /// <summary>Sets a fill where cell text equals the given value. Comparison is case-insensitive by default.</summary>
        public ExcelColumnStyleByHeaderBuilder SetFillColorWhenTextEquals(string text, OfficeColor color, bool caseInsensitive = true) {
            if (string.IsNullOrEmpty(text)) return this;
            var comparison = caseInsensitive ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            return Apply((row, column) => {
                if (_sheet.TryGetCellText(row, column, out var value) && string.Equals(value, text, comparison))
                    _sheet.CellBackground(row, column, color);
            });
        }

        /// <summary>Sets fill colors by matching cell text to a color map.</summary>
        public ExcelColumnStyleByHeaderBuilder SetFillColorByTextMap(IDictionary<string, OfficeColor> map, bool caseInsensitive = true) =>
            ApplyColorMap(map, caseInsensitive, (row, column, color) => _sheet.CellBackground(row, column, color));

        /// <summary>Sets or clears bold where cell text equals the given value.</summary>
        public ExcelColumnStyleByHeaderBuilder SetBoldWhenTextEquals(string text, bool bold = true, bool caseInsensitive = true) {
            if (string.IsNullOrEmpty(text)) return this;
            var comparison = caseInsensitive ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            return Apply((row, column) => {
                if (_sheet.TryGetCellText(row, column, out var value) && string.Equals(value, text, comparison))
                    _sheet.CellBold(row, column, bold);
            });
        }

        /// <summary>Sets or clears bold where cell text is in the given set.</summary>
        public ExcelColumnStyleByHeaderBuilder SetBoldByTextSet(ISet<string> values, bool bold = true, bool caseInsensitive = true) {
            if (values == null || values.Count == 0) return this;
            var set = new HashSet<string>(values, caseInsensitive ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
            return Apply((row, column) => {
                if (_sheet.TryGetCellText(row, column, out var value) && !string.IsNullOrEmpty(value) && set.Contains(value))
                    _sheet.CellBold(row, column, bold);
            });
        }

        private ExcelColumnStyleByHeaderBuilder ApplyColorMap(IDictionary<string, OfficeColor> map, bool caseInsensitive,
            Action<int, int, OfficeColor> apply) {
            if (map == null || map.Count == 0) return this;
            var colors = new Dictionary<string, OfficeColor>(map, caseInsensitive ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
            return Apply((row, column) => {
                if (_sheet.TryGetCellText(row, column, out var value) && !string.IsNullOrEmpty(value) && colors.TryGetValue(value, out var color))
                    apply(row, column, color);
            });
        }

        private ExcelColumnStyleByHeaderBuilder Apply(Action<int, int> apply) {
            if (_colIndex <= 0 || _startRow > _endRow) return this;
            _sheet.RangeEdit(() => {
                for (int row = _startRow; row <= _endRow; row++) apply(row, _colIndex);
            });
            return this;
        }
    }
}

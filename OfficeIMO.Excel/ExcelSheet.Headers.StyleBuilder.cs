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

        public ColumnStyleByHeaderBuilder Number(int decimals = 0)
            => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Decimal, decimals));

        public ColumnStyleByHeaderBuilder Integer() => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Integer));

        public ColumnStyleByHeaderBuilder Percent(int decimals = 0)
            => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Percent, decimals));

        public ColumnStyleByHeaderBuilder Currency(int decimals = 2, CultureInfo? culture = null)
            => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Currency, decimals, culture));

        public ColumnStyleByHeaderBuilder Date(string pattern = "yyyy-mm-dd") => ApplyFormat(pattern);

        public ColumnStyleByHeaderBuilder DateTime(string pattern = "yyyy-mm-dd hh:mm:ss") => ApplyFormat(pattern);

        public ColumnStyleByHeaderBuilder Time() => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Time));

        public ColumnStyleByHeaderBuilder DurationHours() => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.DurationHours));

        public ColumnStyleByHeaderBuilder Text() => ApplyFormat(ExcelNumberFormats.Get(ExcelNumberPreset.Text));

        public ColumnStyleByHeaderBuilder NumberFormat(string format) => ApplyFormat(format);
    }
}


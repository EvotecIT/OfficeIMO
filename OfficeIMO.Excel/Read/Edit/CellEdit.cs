using System;

namespace OfficeIMO.Excel
{
    /// <summary>
    /// Editable cell view used by RowEdit to read current snapshot value and write new value back to the sheet.
    /// </summary>
    public sealed class CellEdit
    {
        private readonly ExcelSheet _sheet;
        private object? _snapshotValue;

        internal CellEdit(ExcelSheet sheet, int rowIndex, int columnIndex, object? snapshotValue)
        {
            _sheet = sheet;
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
            _snapshotValue = snapshotValue;
        }

        /// <summary>
        /// 1-based row index.
        /// </summary>
        public int RowIndex { get; }

        /// <summary>
        /// 1-based column index.
        /// </summary>
        public int ColumnIndex { get; }

        /// <summary>
        /// Current value snapshot from read time.
        /// Setting this property writes directly to the worksheet cell via <see cref="ExcelSheet.CellValue(int,int,object)"/>.
        /// </summary>
        public object? Value
        {
            get => _snapshotValue;
            set
            {
                _snapshotValue = value;
                if (value is null)
                {
                    _sheet.CellValue(RowIndex, ColumnIndex, string.Empty);
                }
                else
                {
                    _sheet.CellValue(RowIndex, ColumnIndex, value);
                }
            }
        }

        /// <summary>
        /// Helper for typed conversion consistent with common Convert semantics.
        /// </summary>
        public T ConvertTo<T>()
        {
            var v = _snapshotValue;
            if (v is null) return default!;
            var dest = typeof(T);
            if (dest.IsAssignableFrom(v.GetType())) return (T)v;
            try
            {
                if (dest == typeof(string)) return (T)(object)(Convert.ToString(v) ?? string.Empty);
                if (dest == typeof(int)) return (T)(object)Convert.ToInt32(v);
                if (dest == typeof(long)) return (T)(object)Convert.ToInt64(v);
                if (dest == typeof(double)) return (T)(object)Convert.ToDouble(v);
                if (dest == typeof(decimal)) return (T)(object)Convert.ToDecimal(v);
                if (dest == typeof(bool)) return (T)(object)Convert.ToBoolean(v);
                if (dest == typeof(DateTime))
                {
                    if (v is double oa) return (T)(object)DateTime.FromOADate(oa);
                    return (T)(object)Convert.ToDateTime(v);
                }
                return (T)Convert.ChangeType(v, dest);
            }
            catch
            {
                return default!;
            }
        }

        /// <summary>
        /// Applies an Excel number format to this cell (e.g., "0.00", "yyyy-mm-dd", "[h]:mm:ss").
        /// </summary>
        public void NumberFormat(string format)
        {
            if (string.IsNullOrWhiteSpace(format)) return;
            _sheet.FormatCell(RowIndex, ColumnIndex, format);
        }

        /// <summary>
        /// Sets a formula on this cell.
        /// </summary>
        public void Formula(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula)) return;
            _sheet.CellFormula(RowIndex, ColumnIndex, formula);
        }
    }
}

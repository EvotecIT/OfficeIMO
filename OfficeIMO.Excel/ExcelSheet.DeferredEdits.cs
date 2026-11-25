using System.Globalization;

namespace OfficeIMO.Excel {
    /// <summary>The value kinds a workbook that has not opened its SDK package records without it.</summary>
    internal enum ExcelDeferredCellKind : byte {
        /// <summary>Text that becomes a shared string, or cell text past the shared-string promotion limit.</summary>
        Text,
        /// <summary>A number already written as invariant text.</summary>
        Number,
        Boolean,
        /// <summary>A date serial written as invariant text, with the automatic date format.</summary>
        Date,
        /// <summary>A style change; the cell keeps its value.</summary>
        Style,
    }

    /// <summary>One cell value written while the SDK package was not open, in the order it was written.</summary>
    internal readonly struct ExcelDeferredCellEdit {
        internal ExcelDeferredCellEdit(ExcelSheet sheet, int row, int column, ExcelDeferredCellKind kind, string text, object value,
            ExcelStyleOperation? style = null, Action? replay = null) {
            Sheet = sheet;
            Row = row;
            Column = column;
            Kind = kind;
            Text = text;
            Value = value;
            Style = style;
            Replay = replay;
        }

        /// <summary>The style change of a <see cref="ExcelDeferredCellKind.Style"/> edit.</summary>
        internal ExcelStyleOperation? Style { get; }
        /// <summary>Applies a style edit through the model when the package opens.</summary>
        internal Action? Replay { get; }

        internal ExcelSheet Sheet { get; }
        internal int Row { get; }
        internal int Column { get; }
        internal ExcelDeferredCellKind Kind { get; }
        /// <summary>The text the cell holds: the string, the invariant number or serial, or "1"/"0".</summary>
        internal string Text { get; }
        /// <summary>The value as the caller passed it, for writing it through the SDK model when the package opens.</summary>
        internal object Value { get; }
    }

    /// <summary>
    /// Plain cell values written to a worksheet whose workbook has not opened its SDK package are recorded here and
    /// written by the owned save (<see cref="ExcelDocument"/>'s deferred save). Anything the recorder does not cover
    /// opens the SDK, which first writes the recorded values through the model in their original order.
    /// </summary>
    public partial class ExcelSheet {
        private bool CanDeferCellValue(int row, int column) =>
            _worksheetPartValue == null && _deferredEntry != null && _excelDocument.IsSdkDeferred
            && row >= 1 && row <= ExcelA1.MaxRows && column >= 1 && column <= ExcelA1.MaxColumns
            && _excelDocument.FileOpenAccess != FileAccess.Read;

        private bool TryDeferCellValue(int row, int column, string? value) {
            if (string.IsNullOrEmpty(value) || !CanDeferCellValue(row, column)) return false;
            // A line break needs the wrap format, which the model applies.
            if (value!.IndexOf('\n') >= 0 || value.IndexOf('\r') >= 0) return false;
            CoerceValueHelper.ValidateSharedStringLength(value, nameof(value));
            return RecordDeferredCell(row, column, ExcelDeferredCellKind.Text, value, value);
        }

        private bool TryDeferCellValue(int row, int column, double value) =>
            !double.IsNaN(value) && !double.IsInfinity(value) && CanDeferCellValue(row, column)
            && RecordDeferredCell(row, column, ExcelDeferredCellKind.Number, InvariantNumberText.Get(value), value);

        private bool TryDeferCellValue(int row, int column, int value) =>
            CanDeferCellValue(row, column)
            && RecordDeferredCell(row, column, ExcelDeferredCellKind.Number, InvariantNumberText.Get(value), value);

        private bool TryDeferCellValue(int row, int column, long value) {
            if (!CanDeferCellValue(row, column)) return false;
            string text = InvariantNumberText.Get(value);
            // Past 15 significant digits the model keeps the exact value as text.
            return ExcelExactNumber.FitsNumberCell(value)
                ? RecordDeferredCell(row, column, ExcelDeferredCellKind.Number, text, value)
                : RecordDeferredCell(row, column, ExcelDeferredCellKind.Text, text, value);
        }

        private bool TryDeferCellValue(int row, int column, decimal value) {
            if (!CanDeferCellValue(row, column)) return false;
            string text = value.ToString(CultureInfo.InvariantCulture);
            return ExcelExactNumber.FitsNumberCell(value)
                ? RecordDeferredCell(row, column, ExcelDeferredCellKind.Number, text, value)
                : RecordDeferredCell(row, column, ExcelDeferredCellKind.Text, text, value);
        }

        private bool TryDeferCellValue(int row, int column, bool value) =>
            CanDeferCellValue(row, column)
            && RecordDeferredCell(row, column, ExcelDeferredCellKind.Boolean, value ? "1" : "0", value);

        private bool TryDeferCellValue(int row, int column, DateTime value) {
            if (!CanDeferCellValue(row, column)) return false;
            double serial = ExcelDateSystemConverter.ToSerial(value, _excelDocument.DateSystem);
            return RecordDeferredCell(row, column, ExcelDeferredCellKind.Date, InvariantNumberText.Get(serial), value);
        }

        private bool TryDeferBoxedCellValue(int row, int column, object? value) => value switch {
            string text => TryDeferCellValue(row, column, text),
            double number => TryDeferCellValue(row, column, number),
            int number => TryDeferCellValue(row, column, number),
            long number => TryDeferCellValue(row, column, number),
            decimal number => TryDeferCellValue(row, column, number),
            bool flag => TryDeferCellValue(row, column, flag),
            DateTime date => TryDeferCellValue(row, column, date),
            _ => false
        };

        private bool RecordDeferredCell(int row, int column, ExcelDeferredCellKind kind, string text, object value) =>
            _excelDocument.RecordDeferredCellEdit(new ExcelDeferredCellEdit(this, row, column, kind, text, value));

        /// <summary>Records a style change for the owned save; false when the model must apply it.</summary>
        private bool TryDeferCellStyle(int row, int column, ExcelStyleOperation operation, Action replay) =>
            CanDeferCellValue(row, column)
            && _excelDocument.RecordDeferredCellEdit(new ExcelDeferredCellEdit(this, row, column, ExcelDeferredCellKind.Style, string.Empty,
                string.Empty, operation, replay));

        private bool TryDeferFontFlag(int row, int column, ExcelStyleOperationKind kind, bool flag, Action replay) =>
            TryDeferCellStyle(row, column, ExcelStyleOperation.Of(kind, flag), replay);

        private bool TryDeferFontName(int row, int column, string fontName) {
            if (!CanDeferCellValue(row, column)) return false;
            ExcelTextValidation.ThrowIfUnpairedSurrogate(fontName, nameof(fontName), "Font name");
            return TryDeferCellStyle(row, column, ExcelStyleOperation.Of(ExcelStyleOperationKind.FontName, fontName), () => CellFontName(row, column, fontName));
        }

        private bool TryDeferNumberFormat(int row, int column, string numberFormat) {
            if (numberFormat == null || !CanDeferCellValue(row, column)) return false;
            ExcelTextValidation.ThrowIfUnpairedSurrogate(numberFormat, nameof(numberFormat), "Number format");
            return TryDeferCellStyle(row, column, ExcelStyleOperation.Of(ExcelStyleOperationKind.NumberFormat, numberFormat), () => FormatCell(row, column, numberFormat));
        }

        /// <summary>Writes a recorded value through the model, as the overload it came from would have written it.</summary>
        internal void ReplayDeferredCellEdit(in ExcelDeferredCellEdit edit) {
            if (edit.Replay != null) {
                edit.Replay();
                return;
            }
            switch (edit.Value) {
                case string text: CellValue(edit.Row, edit.Column, text); break;
                case double number: CellValue(edit.Row, edit.Column, number); break;
                case int number: CellValue(edit.Row, edit.Column, number); break;
                case long number: CellValue(edit.Row, edit.Column, number); break;
                case decimal number: CellValue(edit.Row, edit.Column, number); break;
                case bool flag: CellValue(edit.Row, edit.Column, flag); break;
                case DateTime date: CellValue(edit.Row, edit.Column, date); break;
                default: throw new InvalidOperationException("A recorded cell value has an unexpected type.");
            }
        }
    }
}

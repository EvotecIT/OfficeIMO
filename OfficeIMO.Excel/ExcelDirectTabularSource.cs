namespace OfficeIMO.Excel {
    internal enum ExcelDirectTabularValueKind {
        Empty,
        Text,
        Number,
        Boolean,
        Unsupported
    }

    internal readonly struct ExcelDirectTabularValue {
        private ExcelDirectTabularValue(
            ExcelDirectTabularValueKind kind,
            string? text,
            double number,
            bool boolean) {
            Kind = kind;
            Text = text;
            Number = number;
            Boolean = boolean;
        }

        internal ExcelDirectTabularValueKind Kind { get; }

        internal string? Text { get; }

        internal double Number { get; }

        internal bool Boolean { get; }

        internal static ExcelDirectTabularValue Normalize(object? value) {
            if (value == null || value == DBNull.Value) {
                return new ExcelDirectTabularValue(ExcelDirectTabularValueKind.Empty, null, 0D, false);
            }

            switch (value) {
                case string text:
                    return new ExcelDirectTabularValue(ExcelDirectTabularValueKind.Text, text, 0D, false);
                case bool boolean:
                    return new ExcelDirectTabularValue(ExcelDirectTabularValueKind.Boolean, null, 0D, boolean);
                case byte number:
                    return NumberValue(number);
                case sbyte number:
                    return NumberValue(number);
                case short number:
                    return NumberValue(number);
                case ushort number:
                    return NumberValue(number);
                case int number:
                    return NumberValue(number);
                case uint number:
                    return NumberValue(number);
                case long number:
                    return NumberValue(number);
                case ulong number:
                    return NumberValue(number);
                case float number when !float.IsNaN(number) && !float.IsInfinity(number):
                    return NumberValue(number);
                case double number when !double.IsNaN(number) && !double.IsInfinity(number):
                    return NumberValue(number);
                case decimal number:
                    return NumberValue((double)number);
                default:
                    return new ExcelDirectTabularValue(ExcelDirectTabularValueKind.Unsupported, null, 0D, false);
            }
        }

        private static ExcelDirectTabularValue NumberValue(double value) =>
            new ExcelDirectTabularValue(ExcelDirectTabularValueKind.Number, null, value, false);
    }

    /// <summary>
    /// Describes a validated, worksheet-wide tabular source that native format
    /// writers can consume without expanding it into the Open XML DOM.
    /// </summary>
    internal sealed class ExcelDirectTabularSource {
        internal ExcelDirectTabularSource(
            string sheetName,
            IExcelSheetTabularRowSource rows,
            bool includeHeaders) {
            SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
            Rows = rows ?? throw new ArgumentNullException(nameof(rows));
            IncludeHeaders = includeHeaders;
        }

        internal string SheetName { get; }

        internal IExcelSheetTabularRowSource Rows { get; }

        internal bool IncludeHeaders { get; }
    }
}

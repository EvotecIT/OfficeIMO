namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    public sealed partial class ExcelTabularRowWriter {
        private ExcelTabularRowWriter WriteNativeValue(object? value) => value switch {
            null or DBNull => Write((string?)null),
            string text => Write(text),
            bool boolean => Write(boolean),
            DateTime date => Write(date),
            DateTimeOffset date => Write(date),
            TimeSpan time => Write(time),
            double number => Write(number),
            float number => Write(number),
            decimal number => Write(number),
            byte number => Write(number),
            sbyte number => Write(number),
            short number => Write(number),
            ushort number => Write(number),
            int number => Write(number),
            uint number => Write(number),
            long number => Write(number),
            ulong number => Write(number),
            byte[] bytes => Write(System.Convert.ToBase64String(bytes)),
#if NET6_0_OR_GREATER
            DateOnly date => Write(date.ToDateTime(TimeOnly.MinValue)),
            TimeOnly time => Write(time.ToTimeSpan()),
#endif
            _ => Write(value.ToString() ?? string.Empty)
        };
    }
}

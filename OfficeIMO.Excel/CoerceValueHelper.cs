using System;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

internal static class CoerceValueHelper
{
    internal static (CellValue cellValue, CellValues type) Coerce(object? value, Func<string, CellValue> sharedStringHandler)
    {
        return value switch
        {
            null => (new CellValue(string.Empty), CellValues.String),
            System.DBNull => (new CellValue(string.Empty), CellValues.String),
            string s => HandleSharedString(s),
            double d => (new CellValue(d.ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            float f => (new CellValue(Convert.ToDouble(f).ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            decimal dec => (new CellValue(dec.ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            int i => (new CellValue(((double)i).ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            long l => (new CellValue(((double)l).ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            DateTime dt => (new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            DateTimeOffset dto => (new CellValue(dto.UtcDateTime.ToOADate().ToString(CultureInfo.InvariantCulture)), CellValues.Number),
#if NET6_0_OR_GREATER
            DateOnly dateOnly => (new CellValue(dateOnly.ToDateTime(TimeOnly.MinValue).ToOADate().ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            TimeOnly timeOnly => (new CellValue(timeOnly.ToTimeSpan().TotalDays.ToString(CultureInfo.InvariantCulture)), CellValues.Number),
#endif
            TimeSpan ts => (new CellValue(ts.TotalDays.ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            bool b => (new CellValue(b ? "1" : "0"), CellValues.Boolean),
            uint ui => (new CellValue(((double)ui).ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            ulong ul => (new CellValue(((double)ul).ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            ushort us => (new CellValue(((double)us).ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            byte by => (new CellValue(((double)by).ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            sbyte sb => (new CellValue(((double)sb).ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            short sh => (new CellValue(((double)sh).ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            Guid guid => HandleSharedString(guid.ToString()),
            Enum e => HandleSharedString(e.ToString()),
            char ch => HandleSharedString(ch.ToString()),
            Uri uri => HandleSharedString(uri.ToString()),
            _ => HandleSharedString(value.ToString() ?? string.Empty)
        };

        (CellValue, CellValues) HandleSharedString(string text)
        {
            if (text.Length > 32767)
            {
                throw new ArgumentException("String exceeds Excel's limit of 32,767 characters", nameof(value));
            }
            return (sharedStringHandler(text), CellValues.SharedString);
        }
    }
}

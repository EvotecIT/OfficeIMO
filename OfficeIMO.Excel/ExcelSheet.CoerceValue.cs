using System;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using CellValuesEnum = DocumentFormat.OpenXml.Spreadsheet.CellValues;

namespace OfficeIMO.Excel
{
    public partial class ExcelSheet
    {
        private static (CellValue cellValue, CellValuesEnum dataType) CoerceValueHelper(object value, Func<string, CellValue> sharedStringHandler)
        {
            switch (value)
            {
                case null:
                    return (new CellValue(string.Empty), CellValuesEnum.String);
                case string s:
                    return (sharedStringHandler(s), CellValuesEnum.SharedString);
                case double d:
                    return (new CellValue(d.ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case float f:
                    return (new CellValue(Convert.ToDouble(f).ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case decimal dec:
                    return (new CellValue(dec.ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case int i:
                    return (new CellValue(((double)i).ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case long l:
                    return (new CellValue(((double)l).ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case DateTime dt:
                    return (new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case DateTimeOffset dto:
                    return (new CellValue(dto.UtcDateTime.ToOADate().ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case TimeSpan ts:
                    return (new CellValue(ts.TotalDays.ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case bool b:
                    return (new CellValue(b ? "1" : "0"), CellValuesEnum.Boolean);
                case uint ui:
                    return (new CellValue(((double)ui).ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case ulong ul:
                    return (new CellValue(((double)ul).ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case ushort us:
                    return (new CellValue(((double)us).ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case byte by:
                    return (new CellValue(((double)by).ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case sbyte sb:
                    return (new CellValue(((double)sb).ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case short sh:
                    return (new CellValue(((double)sh).ToString(CultureInfo.InvariantCulture)), CellValuesEnum.Number);
                case Guid guid:
                    return (sharedStringHandler(guid.ToString()), CellValuesEnum.SharedString);
                case Enum e:
                    return (sharedStringHandler(e.ToString()), CellValuesEnum.SharedString);
                case char ch:
                    return (sharedStringHandler(ch.ToString()), CellValuesEnum.SharedString);
                case System.DBNull:
                    return (new CellValue(string.Empty), CellValuesEnum.String);
                case Uri uri:
                    return (sharedStringHandler(uri.ToString()), CellValuesEnum.SharedString);
                default:
                    string stringValue = value?.ToString() ?? string.Empty;
                    return (sharedStringHandler(stringValue), CellValuesEnum.SharedString);
            }
        }
    }
}


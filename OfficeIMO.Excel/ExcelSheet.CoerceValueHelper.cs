using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static (CellValue cellValue, DocumentFormat.OpenXml.Spreadsheet.CellValues type) CoerceValueHelper(object value, Func<string, CellValue> handleSharedString) {
            switch (value) {
                case null:
                    return (new CellValue(string.Empty), DocumentFormat.OpenXml.Spreadsheet.CellValues.String);
                case string s:
                    return (handleSharedString(s), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString);
                case double d:
                    return (new CellValue(d.ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case float f:
                    return (new CellValue(Convert.ToDouble(f).ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case decimal dec:
                    return (new CellValue(dec.ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case int i:
                    return (new CellValue(((double)i).ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case long l:
                    return (new CellValue(((double)l).ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case uint ui:
                    return (new CellValue(((double)ui).ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case ulong ul:
                    return (new CellValue(((double)ul).ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case ushort us:
                    return (new CellValue(((double)us).ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case byte b:
                    return (new CellValue(((double)b).ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case sbyte sb:
                    return (new CellValue(((double)sb).ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case short sh:
                    return (new CellValue(((double)sh).ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case DateTime dt:
                    return (new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case DateTimeOffset dto:
                    return (new CellValue(dto.UtcDateTime.ToOADate().ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case TimeSpan ts:
                    return (new CellValue(ts.TotalDays.ToString(CultureInfo.InvariantCulture)), DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
                case bool bo:
                    return (new CellValue(bo ? "1" : "0"), DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean);
                case Guid guid:
                    var gtext = guid.ToString();
                    return (handleSharedString(gtext), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString);
                case Enum e:
                    var name = e.ToString();
                    return (handleSharedString(name), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString);
                case char ch:
                    var ctext = ch.ToString();
                    return (handleSharedString(ctext), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString);
                case System.DBNull:
                    return (new CellValue(string.Empty), DocumentFormat.OpenXml.Spreadsheet.CellValues.String);
                case Uri uri:
                    var utext = uri.ToString();
                    return (handleSharedString(utext), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString);
                default:
                    var text = value?.ToString() ?? string.Empty;
                    return (handleSharedString(text), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString);
            }
        }
    }
}

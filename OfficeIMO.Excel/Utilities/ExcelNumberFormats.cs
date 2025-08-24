using System;
using System.Globalization;

namespace OfficeIMO.Excel
{
    /// <summary>
    /// Maps number format presets to Excel format codes.
    /// </summary>
    public static class ExcelNumberFormats
    {
        public static string Get(ExcelNumberPreset preset, int decimals = 2, CultureInfo? culture = null)
        {
            culture ??= CultureInfo.CurrentCulture;
            switch (preset)
            {
                case ExcelNumberPreset.General:
                    return "General";
                case ExcelNumberPreset.Integer:
                    return "#,##0";
                case ExcelNumberPreset.Decimal:
                    return "#,##0" + (decimals > 0 ? "." + new string('0', decimals) : string.Empty);
                case ExcelNumberPreset.Percent:
                    return "0" + (decimals > 0 ? "." + new string('0', decimals) : string.Empty) + "%";
                case ExcelNumberPreset.Currency:
                    {
                        var sym = culture.NumberFormat.CurrencySymbol;
                        // Literal currency symbol, basic grouping with configurable decimals
                        return "\"" + sym + "\"#,##0" + (decimals > 0 ? "." + new string('0', decimals) : string.Empty);
                    }
                case ExcelNumberPreset.Scientific:
                    return "0" + (decimals > 0 ? "." + new string('0', decimals) : string.Empty) + "E+00";
                case ExcelNumberPreset.DateShort:
                    return "yyyy-mm-dd";
                case ExcelNumberPreset.DateLong:
                    return "yyyy-mm-dd hh:mm";
                case ExcelNumberPreset.Time:
                    return "h:mm:ss";
                case ExcelNumberPreset.DateTime:
                    return "yyyy-mm-dd hh:mm:ss";
                case ExcelNumberPreset.DurationHours:
                    return "[h]:mm:ss";
                case ExcelNumberPreset.Text:
                    return "@";
                default:
                    return "General";
            }
        }
    }
}

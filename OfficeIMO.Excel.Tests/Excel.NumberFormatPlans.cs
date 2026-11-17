using System.Globalization;
using OfficeIMO.Excel;
using OfficeIMO.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// The compiled number format plan used by displayed-text reads must produce exactly what the per-call display engine
/// produces, so cached plans can never drift from <c>GetFormattedText</c>.
/// </summary>
public sealed class ExcelNumberFormatPlanTests {
    private static readonly (uint Id, string Code)[] Formats = {
        (164, "0.0%"), (10, "0.00%"), (165, "\"$\"#,##0"), (166, "\"$\"#,##0.00"), (167, "#,##0.00;(#,##0.00)"),
        (11, "0.00E+00"), (168, "##0.0E+0"), (12, "# ?/?"), (13, "# ??/??"), (169, "0.###"), (170, "000000"),
        (3, "#,##0"), (171, "#,##0,"), (172, "0;-0;\"zero\""), (173, "\"Total: \"0.0"), (174, "0.00_);[Red](0.00)"),
        (175, "[$€-407] #,##0.00"), (176, "#,##0.00 [$zł-415]"), (177, "#"), (178, "0.0;;"),
        (179, "yyyy-mm-dd"), (14, "mm-dd-yy"), (22, "m/d/yy h:mm"), (180, "m/d/yyyy h:mm AM/PM"), (181, "h:mm:ss.000"),
        (182, "mmmmm"), (183, "dddd, mmmm d, yyyy"), (184, "[$-409]mmm d, yyyy"), (185, "hh:mm A/P"),
        (186, "[>100]0;0.0"), (187, "[h]:mm:ss"), (188, "@\" units\""),
        (189, "General;[Red]General"), (190, "General;(General)"), (191, "\"Total: \"General"), (192, "0.00;-General;\"zero\"")
    };

    private static readonly double[] Values = {
        0, 1, -1, 0.5, -0.5, 0.032, 31200, -1234.5678, 45306.5625, 1e-7, 123456789.123, 0.999999, 99.995, 1e20, -45306.25
    };

    [Theory]
    [InlineData("")]
    [InlineData("de-DE")]
    [InlineData("pl-PL")]
    [InlineData("en-US")]
    public void CompiledPlansMatchThePerCallFormatter(string cultureName) {
        CultureInfo culture = cultureName.Length == 0 ? CultureInfo.InvariantCulture : CultureInfo.GetCultureInfo(cultureName);
        foreach (ExcelDateSystem dateSystem in new[] { ExcelDateSystem.NineteenHundred, ExcelDateSystem.NineteenFour }) {
            foreach ((uint id, string code) in Formats) {
                ExcelNumberFormatPlan plan = ExcelNumberFormatPlan.Create(id, code, dateSystem, culture)!;
                foreach (double value in Values) {
                    string expected = ExcelNumberFormatDisplay.FormatNumericText(value, id, code,
                        SpreadsheetNumberFormatDisplay.FormatGeneralValue(value, culture), dateSystem, culture);
                    Assert.True(expected == plan.Format(value),
                        $"{code} [{cultureName}/{dateSystem}] {value.ToString("R", CultureInfo.InvariantCulture)}: expected '{expected}', got '{plan.Format(value)}'");
                }
            }
        }
    }

    [Theory]
    [InlineData("General;[Red]General", 1234.5, "1234.5")]
    [InlineData("General;[Red]General", -2, "2")]
    [InlineData("General;(General)", -2.25, "(2.25)")]
    [InlineData("\"Total: \"General", 3, "Total: 3")]
    public void GeneralSectionsShowTheValueInGeneralForm(string code, double value, string expected) {
        ExcelNumberFormatPlan plan = ExcelNumberFormatPlan.Create(200, code, ExcelDateSystem.NineteenHundred, CultureInfo.InvariantCulture)!;

        Assert.Equal(expected, plan.Format(value));
    }
}

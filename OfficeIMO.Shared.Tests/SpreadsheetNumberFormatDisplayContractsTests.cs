using System.Globalization;
using OfficeIMO.Spreadsheet;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public class SpreadsheetNumberFormatDisplayContractsTests {
    [Theory]
    [InlineData("nl-NL", "1,23E+03")]
    [InlineData("en-US", "1.23E+03")]
    public void ScientificPayloadUsesProviderWithoutReplacingLiteralPunctuation(string culture, string expected) {
        Assert.Equal(".," + expected + ",.", SpreadsheetNumberFormatDisplay.FormatNumericValue(
            1234d, "\".,\"0.00E+00\",.\"", CultureInfo.GetCultureInfo(culture)));
    }

    [Theory]
    [InlineData("0.0#", "nl-NL", "12,5")]
    [InlineData("000.0#", "nl-NL", "012,5")]
    [InlineData("#,##0.00", "en-US", "12.50")]
    public void OptionalPrecisionAndIntegerPaddingUseProviderSeparators(string format, string culture, string expected) {
        Assert.Equal(expected, SpreadsheetNumberFormatDisplay.FormatNumericValue(12.5, format, CultureInfo.GetCultureInfo(culture)));
    }

    [Theory]
    [InlineData("0;0;0;[Unknown]@", "[Unknown]abc")]
    [InlineData("0;0;0;[ColorUnknown]@", "[ColorUnknown]abc")]
    [InlineData("0;0;0;\"@;\"@", "@;abc")]
    public void TextSectionRetainsUnknownSyntaxAndQuotedPlaceholders(string format, string expected) {
        Assert.Equal(expected, SpreadsheetNumberFormatDisplay.FormatTextValue("abc", format));
    }

    [Theory]
    [InlineData(0d, "-")]
    [InlineData(2d, "2.00")]
    public void AccountingZeroSectionPrintsItsDashWithoutDigits(double value, string expected) {
        Assert.Equal(expected, SpreadsheetNumberFormatDisplay.FormatNumericValue(
            value, "0.00;-0.00;_(* \"-\"??_)", CultureInfo.InvariantCulture).Trim());
    }

    [Theory]
    [InlineData(1.005, "0.00", "1.01")]
    [InlineData(2.675, "0.00", "2.68")]
    [InlineData(-1.005, "0.00", "-1.01")]
    [InlineData(1.004, "0.00", "1.00")]
    public void DecimalMidpointsRoundAwayFromZeroAsSpreadsheetsShowThem(double value, string format, string expected) {
        Assert.Equal(expected, SpreadsheetNumberFormatDisplay.FormatNumericValue(value, format, CultureInfo.InvariantCulture));
    }

    [Fact]
    public void FractionalSecondTimeFormatsStayTimes() {
        Assert.Equal("12:00:00.00", SpreadsheetNumberFormatDisplay.FormatSerialValue(0.5, "hh:mm:ss.00"));
        Assert.Equal("0.50", SpreadsheetNumberFormatDisplay.FormatSerialValue(0.5, "0.00"));
    }

    [Theory]
    [InlineData(-1d)]
    [InlineData(2958465.9999999995)]
    public void DateSerialsOutsideTheDateSystemShowTheInvalidDateMarker(double serial) {
        Assert.Equal("#####", SpreadsheetNumberFormatDisplay.FormatSerialValue(serial, "yyyy-mm-dd"));
    }

    [Fact]
    public void TheLastSerialOfTheDateSystemIsStillADate() {
        Assert.Equal("9999-12-31", SpreadsheetNumberFormatDisplay.FormatSerialValue(2958465d, "yyyy-mm-dd"));
    }
}

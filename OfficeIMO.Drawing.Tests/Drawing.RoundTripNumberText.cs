using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Json;
using Xunit;

namespace OfficeIMO.Tests;

// Numbers written at full precision use the shortest text that parses back, laid out as .NET Core's "R" lays it
// out, on every target framework. The expected texts are .NET Core's; the values are compiler literals, so they do
// not depend on the runtime's parser.
public sealed class DrawingRoundTripNumberTextTests {
    private const double SixteenDigits = 0.4694715627858908D;
    private const string SixteenDigitText = "0.4694715627858908";

    [Theory]
    // .NET Framework's "R" wrote 17 digits for these.
    [InlineData(0.4694715627858908D, "0.4694715627858908")]
    [InlineData(512.0000152587891D, "512.0000152587891")]
    // .NET Framework's parser accepted a shorter text that a correct parser reads as a neighbouring double.
    [InlineData(-7.843380482104101E-145D, "-7.843380482104101E-145")]
    [InlineData(3.7998716159534622E-84D, "3.7998716159534622E-84")]
    [InlineData(5.5028128696153034E-06D, "5.5028128696153034E-06")]
    [InlineData(1.0781753810392959E+194D, "1.0781753810392959E+194")]
    // Its 16-digit rounding picked a different last digit.
    [InlineData(0.9841210986413625D, "0.9841210986413625")]
    [InlineData(802.4644000037252D, "802.4644000037252")]
    // Two decimals are equally close; the even one is written.
    [InlineData(-2714965.6274414062D, "-2714965.6274414062")]
    [InlineData(713731179413504.2D, "713731179413504.2")]
    // Layout: fixed notation for leading exponents from -4 to 16.
    [InlineData(1E-05D, "1E-05")]
    [InlineData(0.0001D, "0.0001")]
    [InlineData(1E+16D, "10000000000000000")]
    [InlineData(1E+17D, "1E+17")]
    [InlineData(123456789012345680D, "1.2345678901234568E+17")]
    [InlineData(5E-324D, "5E-324")]
    [InlineData(-0D, "-0")]
    public void DoublesUseDotNetCoreRoundTripText(double value, string expected) {
        Assert.Equal(expected, OfficeRoundTripText.Format(value, CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData(1F / 7F, "0.14285715")]
    [InlineData(-2321655.2F, "-2321655.2")]
    [InlineData(294614.62F, "294614.62")]
    [InlineData(123456790F, "123456790")]
    [InlineData(1E+09F, "1E+09")]
    [InlineData(1E-05F, "1E-05")]
    [InlineData(3.4028235E+38F, "3.4028235E+38")]
    // Powers of two, where the gap to the next lower float is half as wide, and a value whose shorter probes carry.
    [InlineData(1.5474251E+26F, "1.5474251E+26")]
    [InlineData(1.2621775E-29F, "1.2621775E-29")]
    [InlineData(0.9999999F, "0.9999999")]
    public void FloatsUseDotNetCoreRoundTripText(float value, string expected) {
        Assert.Equal(expected, OfficeRoundTripText.Format(value, CultureInfo.InvariantCulture));
    }

    // .NET Framework's parser reads .NET Core's shortest text for these as a neighbouring double, so on that runtime
    // they get the 17-digit text instead. Either way the text reads back with the runtime's own parser.
    [Theory]
    [InlineData(0.91241918127630794D, "0.912419181276308", "0.91241918127630794")]
    [InlineData(3.2506899999999998E-06D, "3.25069E-06", "3.2506899999999998E-06")]
    [InlineData(-3.0597645285442957E-48D, "-3.059764528544296E-48", "-3.0597645285442957E-48")]
    [InlineData(4.4465908125712189E-323D, "4.4E-323", "4.4465908125712189E-323")]
    public void DoublesReadBackWithTheRuntimeParser(double value, string shortest, string framework) {
        string text = OfficeRoundTripText.Format(value, CultureInfo.InvariantCulture);
#if NETFRAMEWORK
        Assert.Equal(framework, text);
#else
        Assert.Equal(shortest, text);
#endif
        Assert.Equal(BitConverter.DoubleToInt64Bits(value),
            BitConverter.DoubleToInt64Bits(double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture)));
    }

    [Fact]
    public void RoundTripTextUsesTheCultureSeparators() {
        CultureInfo german = CultureInfo.GetCultureInfo("de-DE");
        Assert.Equal("1234,5", OfficeRoundTripText.Format(1234.5D, german));
        Assert.Equal("-1,5E-20", OfficeRoundTripText.Format(-1.5E-20D, german));
    }

    [Fact]
    public void JsonNumbersKeepTheShortestRoundTripText() {
        Assert.Equal(SixteenDigitText, JsonData.CreateNumber(SixteenDigits).ToString());
    }

    [Fact]
    public void PortableJsonWritesTheShortestRoundTripText() {
        OfficeDrawing drawing = new OfficeDrawing(100, 100).AddShape(OfficeShape.Rectangle(40, 20), SixteenDigits, 7);
        ((OfficeDrawingShape)drawing.Elements[0]).Shape.FillColor = OfficeColor.Black;

        Assert.True(OfficeDrawingPortableExporter.TryExportJson(drawing, new OfficeDrawingPortableOptions("OfficeIMO Baseline Sans"),
            out string? json, out _));
        Assert.Contains(SixteenDigitText, json);
        Assert.DoesNotContain(SixteenDigitText + "1", json);
    }

    [Fact]
    public void EditablePathDataWritesTheShortestRoundTripText() {
        string path = OfficeSvgFormatting.FormatEditablePathData(new[] {
            OfficePathCommand.MoveTo(SixteenDigits, 1D), OfficePathCommand.LineTo(2D, 3D) });

        Assert.Equal("M" + SixteenDigitText + " 1L2 3", path);
    }

    [Fact]
    public void RadialPatternBoundsWriteTheShortestRoundTripText() {
        var builder = new StringBuilder();
        builder.AppendNativeRadialPatternDefinition("p", SixteenDigits, 0D, 10D, 10D, "#000000", "1",
            (_, _) => { }, (_, _) => { });

        Assert.Contains("x=\"" + SixteenDigitText + "\"", builder.ToString());
    }

    [Fact]
    public void OutlinedTextWritesTheShortestFillOpacity() {
        var drawing = new OfficeDrawing(100, 40).AddText("A", 0, 0, 100, 40, color: OfficeColor.FromRgba(0, 0, 0, 128));
        var text = (OfficeDrawingText)drawing.Elements[0];
        text.OutlineStrokeColor = OfficeColor.Black;
        text.OutlineStrokeWidth = 1D;

        // 128 / 255 needs 16 digits.
        Assert.Contains("fill-opacity=\"0.5019607843137255\"", OfficeDrawingSvgExporter.ToSvg(drawing));
    }
}

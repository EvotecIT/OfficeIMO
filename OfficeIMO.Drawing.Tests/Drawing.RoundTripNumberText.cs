using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Json;
using Xunit;

namespace OfficeIMO.Tests;

// Numbers written at full precision use the shortest text that parses back, on every target framework.
// .NET Framework's "R" format jumps from 15 to 17 digits (9 for a float), so these values, which need
// 16 digits (8 for a float), came out longer there than on .NET 8 and 10.
public sealed class DrawingRoundTripNumberTextTests {
    private const double SixteenDigits = 0.4694715627858908D;
    private const string SixteenDigitText = "0.4694715627858908";

    [Fact]
    public void FloatsUseTheirShortestRoundTripText() {
        Assert.Equal("0.14285715", OfficeRoundTripText.Format(1F / 7F, CultureInfo.InvariantCulture));
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

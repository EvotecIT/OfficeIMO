using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeDrawingSvgExporter_LeavesBrowserNativeWidthUnconstrainedWithoutExplicitAdvance() {
        var drawing = new OfficeDrawing(240D, 60D)
            .AddText("Browser shaped", 10D, 10D, 220D, 40D,
                new OfficeFontInfo("Host-only font", 18D), wrapText: false);

        XElement text = Assert.Single(XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing))
            .Descendants(), element => element.Name.LocalName == "text");

        Assert.Equal("browser-native", text.Attribute("data-officeimo-shaping-backend")?.Value);
        Assert.Null(text.Attribute("textLength"));
        Assert.Null(text.Attribute("lengthAdjust"));
    }

    [Theory]
    [InlineData(OfficeTextAlignment.Left, "10", "end")]
    [InlineData(OfficeTextAlignment.Center, "120", "middle")]
    [InlineData(OfficeTextAlignment.Right, "230", "start")]
    public void OfficeDrawingSvgExporter_PreservesPhysicalAlignmentForRightToLeftText(
        OfficeTextAlignment alignment,
        string expectedX,
        string expectedAnchor) {
        var drawing = new OfficeDrawing(240D, 60D)
            .AddText("שלום", 10D, 10D, 220D, 40D,
                new OfficeFontInfo("Arial", 18D), alignment: alignment, wrapText: false);

        XElement text = Assert.Single(XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing))
            .Descendants(), element => element.Name.LocalName == "text");

        Assert.Equal("rtl", text.Attribute("direction")?.Value);
        Assert.Equal(expectedX, text.Attribute("x")?.Value);
        Assert.Equal(expectedAnchor, text.Attribute("text-anchor")?.Value);
    }

    [Fact]
    public void OfficeSvgDrawingReader_PreservesPhysicalLeftAlignmentForRightToLeftText() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 240 60'>"
            + "<text x='10' y='30' font-family='Arial' font-size='18' fill='black' "
            + "direction='rtl' unicode-bidi='plaintext' text-anchor='end'>שלום</text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            System.Text.Encoding.UTF8.GetBytes(svg), options: null, out OfficeDrawing? imported, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeDrawingText text = Assert.Single(imported!.Elements.OfType<OfficeDrawingText>());

        Assert.Equal(10D, text.X, precision: 3);
        Assert.True(text.Width > 20D, $"Expected measured RTL width, got {text.Width}.");
        Assert.Null(text.TextAdvanceWidth);
        Assert.Equal(OfficeTextDirection.Auto, text.TextDirection);
    }

    [Fact]
    public void OfficeSvgDrawingReader_PlaintextUsesFirstStrongDirectionOverAuthoredDirection() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 240 60'>"
            + "<text x='10' y='30' font-family='Arial' font-size='18' fill='black' "
            + "direction='ltr' unicode-bidi='plaintext' text-anchor='end'>שלום</text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            System.Text.Encoding.UTF8.GetBytes(svg), options: null, out OfficeDrawing? imported, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeDrawingText text = Assert.Single(imported!.Elements.OfType<OfficeDrawingText>());

        Assert.Equal(10D, text.X, precision: 3);
        Assert.True(text.Width > 20D, $"Expected measured plaintext RTL width, got {text.Width}.");
        Assert.Equal(OfficeTextDirection.Auto, text.TextDirection);
    }

    [Fact]
    public void OfficeSvgDrawingReader_DirectionUnsetRetainsInheritedDirection() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 240 60'>"
            + "<g direction='rtl'><text x='10' y='30' font-family='Arial' font-size='18' fill='black' "
            + "style='direction: unset' text-anchor='end'>שלום</text></g></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            System.Text.Encoding.UTF8.GetBytes(svg), options: null, out OfficeDrawing? imported, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeDrawingText text = Assert.Single(imported!.Elements.OfType<OfficeDrawingText>());

        Assert.Equal(10D, text.X, precision: 3);
        Assert.True(text.Width > 20D, $"Expected measured inherited RTL width, got {text.Width}.");
        Assert.Equal(OfficeTextDirection.RightToLeft, text.TextDirection);
    }

    [Fact]
    public void OfficeSvgDrawingReader_PreservesAuthoredDirectionWhenFirstStrongTextDiffers() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 320 60'>"
            + "<text x='300' y='30' font-family='Arial' font-size='18' fill='black' "
            + "direction='rtl' text-anchor='start'>abc אבג</text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            System.Text.Encoding.UTF8.GetBytes(svg), options: null, out OfficeDrawing? imported, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeDrawingText text = Assert.Single(imported!.Elements.OfType<OfficeDrawingText>());

        Assert.Equal(OfficeTextDirection.RightToLeft, text.TextDirection);
        XElement exported = Assert.Single(XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(imported))
            .Descendants(), element => element.Name.LocalName == "text");
        Assert.Equal("rtl", exported.Attribute("direction")?.Value);
        Assert.Null(exported.Attribute("unicode-bidi"));
    }

    [Fact]
    public void OfficeDrawingSvgExporter_PreservesVerticalTextDecorations() {
        var drawing = new OfficeDrawing(120D, 180D)
            .AddVerticalText(
                "Vertical",
                20D,
                10D,
                80D,
                160D,
                new OfficeFontInfo("Arial", 24D, OfficeFontStyle.Underline | OfficeFontStyle.Strikethrough),
                OfficeColor.FromRgb(12, 34, 56));

        XElement text = Assert.Single(XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing))
            .Descendants(), element => element.Name.LocalName == "text");

        Assert.Equal("underline line-through", text.Attribute("text-decoration")?.Value);
        Assert.Equal("#0C2238", text.Attribute("text-decoration-color")?.Value);
    }

    [Fact]
    public void OfficeSvgDrawingReader_PlacesRightToLeftTspanRunsInVisualOrder() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 240 60'>"
            + "<text x='200' y='30' font-family='Arial' font-size='18' fill='black' "
            + "direction='rtl' text-anchor='start'><tspan>אב</tspan><tspan>גד</tspan></text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            System.Text.Encoding.UTF8.GetBytes(svg), options: null, out OfficeDrawing? imported, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeDrawingText[] runs = imported!.Elements.OfType<OfficeDrawingText>().ToArray();

        Assert.Equal(2, runs.Length);
        Assert.Equal("אב", runs[0].Text);
        Assert.Equal("גד", runs[1].Text);
        Assert.True(runs[0].X > runs[1].X,
            $"Expected the first logical RTL run at the visual right, got {runs[0].X} and {runs[1].X}.");
        Assert.All(runs, run => Assert.Equal(OfficeTextDirection.RightToLeft, run.TextDirection));
    }

    [Fact]
    public void OfficeSvgDrawingReader_PlacesRightToLeftPositionedGlyphRunsInVisualOrder() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 240 60'>"
            + "<text x='200' y='30' rotate='0 0 0 0' font-family='Arial' font-size='18' fill='black' "
            + "direction='rtl' text-anchor='start'>אבגד</text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            System.Text.Encoding.UTF8.GetBytes(svg), options: null, out OfficeDrawing? imported, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeDrawingText[] runs = imported!.Elements.OfType<OfficeDrawingText>().ToArray();

        Assert.Equal(4, runs.Length);
        Assert.True(runs.Zip(runs.Skip(1), static (left, right) => left.X > right.X).All(static descending => descending),
            "Expected logical RTL glyph runs to descend across the physical X axis.");
        Assert.All(runs, run => Assert.Equal(OfficeTextDirection.RightToLeft, run.TextDirection));
    }
}

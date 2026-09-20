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
}

using System;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsStrokeOutlineTests {
    [Theory]
    [InlineData(XpsFormat.Xps, "Triangle", false)]
    [InlineData(XpsFormat.OpenXps, "Square", false)]
    [InlineData(XpsFormat.Xps, "Triangle", true)]
    [InlineData(XpsFormat.OpenXps, "Square", true)]
    public void DegenerateVerticesDoNotAddUnorientedZeroDashCaps(XpsFormat format, string dashCap, bool afterGap) {
        const string path = "M40,40L112,136";
        const string leading = "M40,40L40,40L112,136";
        const string intermediate = "M40,40L64,72L64,72L112,136";
        var expected = Render(Create(format, path, "Flat", "Flat", dashCap, "0 1"));
        var actual = Render(Create(format, afterGap ? intermediate : leading, "Flat", "Flat", dashCap, "0 1"));
        long difference = 0;
        for (int y = 0; y < expected.Height; y++) for (int x = 0; x < expected.Width; x++)
            difference += Math.Abs(expected.GetPixel(x,y).R - actual.GetPixel(x,y).R);
        Assert.InRange(difference / (double)(expected.Width * expected.Height), 0D, .05D);
    }

    [Theory]
    [InlineData(XpsFormat.Xps, false)]
    [InlineData(XpsFormat.OpenXps, true)]
    public void PaintedDashesRetainAuthoredDegenerateJoins(XpsFormat format, bool explicitGeometry) {
        var doc = Create(format, "M20,70L70,70L70,70L70,20", "Flat", "Flat", "Flat", "10 1");
        if (explicitGeometry) {
            var page = doc.Pages[0]; var xml = page.GetMarkup(); var ns = xml.Name.Namespace; var path = xml.Elements().Single();
            path.Attribute("Data")!.Remove();
            path.Add(new XElement(ns+"Path.Data", new XElement(ns+"PathGeometry", new XElement(ns+"PathFigure",
                new XAttribute("StartPoint","20,70"), new XElement(ns+"PolyLineSegment",new XAttribute("Points","70,70 70,70 70,20"))))));
            page.ReplaceMarkup(xml);
        }
        var image = Render(doc);
        Assert.True(image.GetPixel(76,76).R < 30);
        Assert.True(image.GetPixel(78,78).R > 240);
    }

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void PaintedClosedSeamsRetainLeadingDegenerateSegments(XpsFormat format) {
        var image = Render(Create(format, "M70,70L70,70L70,20L20,20L20,70Z", "Flat", "Flat", "Flat", "20 1"));
        Assert.True(image.GetPixel(76,76).R < 30);
        Assert.True(image.GetPixel(78,78).R > 240);
    }

    [Theory]
    [InlineData(XpsFormat.Xps, null)]
    [InlineData(XpsFormat.OpenXps, "10 1")]
    public void OpenReturningFiguresRetainTheirEndpointCaps(XpsFormat format, string? dashes) {
        var doc = Create(format, "M80,80L140,80L80,80", "Triangle", "Round", "Flat", dashes);
        var image = Render(doc);
        Assert.True(image.GetPixel(73,85).R < 30);
        Assert.True(image.GetPixel(69,80).R > 240);
    }

    [Theory]
    [InlineData(XpsFormat.Xps, null, false)]
    [InlineData(XpsFormat.OpenXps, "1 .5", false)]
    [InlineData(XpsFormat.Xps, null, true)]
    [InlineData(XpsFormat.OpenXps, "1 .5", true)]
    public void MagnifiedShortSegmentsKeepTheirTangentAndDashCoverage(XpsFormat format, string? dashes, bool arc) {
        var expected = Render(Create(format, arc ? "M80,80A20,20 0 0 1 80.0004,80.0008" : "M80,80L80.0004,80.0008", "Triangle", "Round", "Flat", dashes));
        var doc = Create(format, arc ? "M.08,.08A.02,.02 0 0 1 .0800004,.0800008" : "M.08,.08L.0800004,.0800008", "Triangle", "Round", "Flat", dashes);
        var page = doc.Pages[0]; var xml = page.GetMarkup(); var path = xml.Elements().Single();
        path.SetAttributeValue("StrokeThickness", ".02"); path.SetAttributeValue("RenderTransform", "1000,0,0,1000,0,0"); page.ReplaceMarkup(xml);
        var actual = Render(doc); long difference = 0;
        for (int y = 0; y < expected.Height; y++) for (int x = 0; x < expected.Width; x++)
            difference += Math.Abs(expected.GetPixel(x,y).R - actual.GetPixel(x,y).R);
        Assert.InRange(difference / (double)(expected.Width * expected.Height), 0D, .05D);
    }

    [Theory]
    [InlineData("M80,80L80,80Z")]
    [InlineData("M80,80Z")]
    public void ClosedDegenerateFiguresDrawTheSpecifiedDot(string geometry) {
        var doc = Create(XpsFormat.OpenXps, geometry, "Flat", "Flat", "Flat");
        var image = Render(doc);
        Assert.True(image.GetPixel(80,80).R < 30);
        Assert.True(image.GetPixel(85,85).R < 30);
        Assert.True(image.GetPixel(88,88).R > 240);
    }

    [Fact]
    public void BrushStrokeCoverageDoesNotInheritTheFillOpacity() {
        var doc = Create(XpsFormat.OpenXps, "M40,80L160,80", "Triangle", "Round", "Flat");
        doc.AddResource("Resources/brush.png", OfficePngWriter.Encode(new OfficeRasterImage(1,1,OfficeColor.Blue)), "image/png");
        var page = doc.Pages[0]; var xml = page.GetMarkup(); var ns = xml.Name.Namespace; var path = xml.Elements().Single();
        path.SetAttributeValue("Fill", "#80FF0000"); path.Attribute("Stroke")!.Remove();
        path.Add(new XElement(ns+"Path.Stroke", new XElement(ns+"ImageBrush", new XAttribute("ImageSource","/Resources/brush.png"),
            new XAttribute("Viewbox","0,0,1,1"), new XAttribute("Viewport","30,70,145,20"), new XAttribute("Opacity",".5"))));
        page.ReplaceMarkup(xml); var image = Render(doc);
        Assert.InRange(image.GetPixel(45,80).R,(byte)126,(byte)129);
        Assert.True(image.GetPixel(45,80).B > 240);
    }

    [Fact]
    public void LargeNativeTransformsRetainEquivalentCurvesAndRoundDashCaps() {
        var original = Create(XpsFormat.Xps, "M20,120C70,0 110,0 170,120", "Triangle", "Square", "Round", "1 .5");
        var scaled = Create(XpsFormat.Xps, "M.02,.12C.07,0 .11,0 .17,.12", "Triangle", "Square", "Round", "1 .5");
        var page = scaled.Pages[0]; var xml = page.GetMarkup(); var path = xml.Elements().Single();
        path.SetAttributeValue("StrokeThickness", ".02"); path.SetAttributeValue("RenderTransform", "1000,0,0,1000,0,0"); page.ReplaceMarkup(xml);
        var expected = Render(original); var actual = Render(scaled);
        long difference = 0;
        for (int y = 0; y < expected.Height; y++) for (int x = 0; x < expected.Width; x++)
            difference += Math.Abs(expected.GetPixel(x,y).R - actual.GetPixel(x,y).R);
        Assert.InRange(difference / (double)(expected.Width * expected.Height), 0D, .05D);
    }

    [Fact]
    public void NativeStrokeExpansionHasAnEnforcedResourceLimit() {
        var doc = Create(XpsFormat.Xps, "M40,80L160,80", "Triangle", "Round", "Round", ".00000001 .00000001");
        Assert.Throws<System.IO.InvalidDataException>(() => doc.Pages[0].ToSvg());
    }

    [Theory]
    [InlineData(XpsFormat.Xps, false)]
    [InlineData(XpsFormat.OpenXps, true)]
    public void ClippedMitersRetainTheBisectorCutAndDegenerateLimit(XpsFormat format, bool degenerate) {
        var doc = Create(format, degenerate ? "M20,70L70,70L70,70L70,20" : "M20,70L70,70L70,20",
            "Flat", "Flat", "Flat", null, degenerate ? 10 : 1);
        var image = Render(doc);
        Assert.True(image.GetPixel(76,76).B > 200 && image.GetPixel(76,76).R < 30);
        Assert.True(image.GetPixel(80,80).R > 240);
        Assert.True(doc.Pages[0].ToSvg().IsComplete);
        Assert.DoesNotContain("/Subtype /Image", System.Text.Encoding.ASCII.GetString(doc.ToPdf()));
    }

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void TriangleStartAndRoundEndHaveDifferentCoverage(XpsFormat format) {
        var doc = Create(format, "M40,80L160,80", "Triangle", "Round", "Flat");
        var image = Render(doc);
        Assert.True(image.GetPixel(35,80).R < 30);
        Assert.True(image.GetPixel(35,86).R > 240);
        Assert.True(image.GetPixel(165,86).R < 30);
        var reimport = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(doc.ToPdf()).Pages[0].ToDrawing(), scale: 4D/3D, background: OfficeColor.White);
        Assert.True(reimport.GetPixel(35,80).R < 30);
        Assert.True(reimport.GetPixel(35,86).R > 240);
    }

    [Theory]
    [InlineData("Flat", 45, false)]
    [InlineData("Round", 45, true)]
    [InlineData("Triangle", 45, true)]
    public void DashCapsAndFigureCapsRemainIndependent(string dashCap, int x, bool painted) {
        var doc = Create(XpsFormat.OpenXps, "M40,80L160,80", "Square", "Flat", dashCap, "0,1");
        var image = Render(doc);
        Assert.Equal(painted, image.GetPixel(x,80).R < 100);
        Assert.True(image.GetPixel(35,86).R < 30); // authored square start cap
    }

    [Fact]
    public void OverlappingCapsPaintTheUnionOnce() {
        var doc = Create(XpsFormat.Xps, "M40,80L160,80", "Round", "Triangle", "Round", "1,.25", color: "#800000FF");
        var image = Render(doc);
        Assert.InRange(image.GetPixel(61,80).R, (byte)126, (byte)128);
        Assert.InRange(image.GetPixel(64,80).R, (byte)126, (byte)128);
    }

    private static OfficeRasterImage Render(XpsDocument doc) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(doc.Pages[0].ExportImage(OfficeImageExportFormat.Png).Bytes, out var image));
        return image!;
    }

    internal static XpsDocument Create(XpsFormat format, string geometry, string startCap, string endCap,
        string dashCap, string? dashes = null, double miterLimit = 10, string color = "#FF0000FF") {
        var doc = XpsDocument.Create(format); var page = doc.AddPage(200,160).AddPath(geometry, null, color, 20);
        var xml = page.GetMarkup(); var path = xml.Elements().Single();
        path.SetAttributeValue("StrokeStartLineCap", startCap); path.SetAttributeValue("StrokeEndLineCap", endCap);
        path.SetAttributeValue("StrokeDashCap", dashCap); path.SetAttributeValue("StrokeMiterLimit", miterLimit);
        if (dashes != null) path.SetAttributeValue("StrokeDashArray", dashes);
        page.ReplaceMarkup(xml); return doc;
    }
}

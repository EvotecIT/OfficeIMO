using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgLineLabelTests {
    [Theory]
    [InlineData(false, 30)]
    [InlineData(true, 30)]
    [InlineData(false, -30)]
    [InlineData(true, -30)]
    [InlineData(false, 90)]
    [InlineData(true, 90)]
    [InlineData(false, 180)]
    [InlineData(true, 180)]
    public void RotatedStraightCaptionsUseTransformedEndpointsWithoutScalingTheirFonts(bool line, double degrees) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = line ? page.Shapes.AddLine(OdfLength.Points(80), OdfLength.Points(40), OdfLength.Points(180), OdfLength.Points(40)) :
            page.Shapes.AddConnector(new OfficePoint(80, 40), new OfficePoint(180, 40));
        AddLabel(document, shape);
        shape.Transform = "rotate(" + (degrees * Math.PI / 180).ToString("R", System.Globalization.CultureInfo.InvariantCulture) + ") translate(220pt 250pt)";
        var transform = OfficeTransform.RotateDegrees(-degrees).Then(OfficeTransform.Translate(220, 250));
        var expected = transform.TransformPoint(new OfficePoint(130, 40));
        string content = document.GetXml("content.xml").ToString();
        foreach (var read in RoundTrips(document)) {
            var projected = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var frame = Assert.Single(Frames(projected.Value));
            Assert.Equal(expected.X, Center(frame).X, 6); Assert.Equal(expected.Y, Center(frame).Y, 6);
            var directed = line ? OfficeTransform.RotateDegrees(-degrees) : OfficeTransform.Identity;
            Assert.Equal(directed.M11, frame.Transform.M11, 6); Assert.Equal(directed.M12, frame.Transform.M12, 6);
            Assert.Equal("Primary emphasized\nSecond line", frame.Text.PlainText);
            Assert.All(frame.Text.Paragraphs.SelectMany(p => p.Runs), run => Assert.Equal(12, run.FontSize));
            Assert.Contains("Second line", OfficeDrawingSvgExporter.ToSvg(projected.Value));
            Assert.NotEmpty(OfficeDrawingRasterRenderer.ToPng(projected.Value));
        }
        Assert.Equal(content, document.GetXml("content.xml").ToString());
    }

    [Theory]
    [InlineData("scale(1.2 1.2)")]
    [InlineData("scale(-1 1) translate(300pt 0pt)")]
    [InlineData("skewX(0.2)")]
    [InlineData("matrix(1 0 0 0 0pt 0pt)")]
    [InlineData("attached-rotation")]
    public void UnqualifiedCaptionTransformsRetainTheirNativeXmlAndGeometryWithAnExplicitLoss(string transform) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddConnector(new OfficePoint(40, 40), new OfficePoint(160, 80));
        shape.Name = "Transform boundary"; AddLabel(document, shape);
        if (transform == "attached-rotation") {
            shape.AttachStartToShape(page.Shapes.AddRectangle(new OdfRect(OdfLength.Points(20), OdfLength.Points(20),
                OdfLength.Points(40), OdfLength.Points(40))));
            shape.Transform = "rotate(0.2)";
        } else shape.Transform = transform;
        string content = document.GetXml("content.xml").ToString();
        var projected = page.ToDrawing();
        Assert.Empty(Frames(projected.Value));
        Assert.Contains(projected.Report.Mappings, m => m.Feature == "shape:Transform boundary:text" && m.Status == OdfConversionMappingStatus.Skipped);
        Assert.Contains("stroke", OfficeDrawingSvgExporter.ToSvg(projected.Value));
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(content, document.GetXml("content.xml").ToString());
    }
}

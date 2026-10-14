using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioRasterRotationTests {
    [Theory]
    [InlineData(30, false, false)]
    [InlineData(-30, false, false)]
    [InlineData(30, true, false)]
    [InlineData(-30, true, false)]
    [InlineData(30, false, true)]
    [InlineData(-30, false, true)]
    [InlineData(30, true, true)]
    [InlineData(-30, true, true)]
    public void NativePlainAndRichTextBoxesRotateInThePageDirection(double degrees, bool connector, bool rich) {
        var document = VisioDocument.Create(); var page = document.AddPage("Rotation", 8, 8);
        var style = new VisioTextStyle { Size = 18, BackgroundColor = OfficeColor.Red, TextAngle = degrees * Math.PI / 180 };
        if (connector) {
            var line = page.AddConnector("c", new OfficePoint(1, 2), new OfficePoint(7, 2));
            line.LinePattern = 0; line.Label = "MMMMMMMM"; line.TextStyle = style;
            line.PlaceLabelAt(4, 4, 3, .6);
        } else {
            var shape = page.AddRectangle(4, 4, 3, .6, "MMMMMMMM", VisioMeasurementUnit.Inches);
            shape.FillPattern = 0; shape.LinePattern = 0; shape.TextStyle = style;
        }
        if (rich) {
            XNamespace ns = "http://schemas.microsoft.com/visio/2003/core";
            var xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
            var shape = xml.Descendants(ns + "Page").Descendants(ns + "Shape").Single();
            var row = new XElement(shape.Element(ns + "Char")!); row.SetAttributeValue("IX", "1");
            row.SetElementValue(ns + "Style", "1"); shape.Add(row);
            shape.Element(ns + "Text")!.ReplaceNodes(new XElement(ns + "cp", new XAttribute("IX", "0")), "MMMM", new XElement(ns + "cp", new XAttribute("IX", "1")), "MMMM");
            document = VisioDocument.LoadLegacyXml(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml.ToString()))).Value;
        }
        foreach (var candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) })
            AssertRedOrientation(candidate.Pages[0], degrees);
    }

    [Theory]
    [InlineData(30)]
    [InlineData(-30)]
    public void EllipseAxesRotateInThePageDirection(double degrees) {
        var page = VisioDocument.Create().AddPage("Ellipse", 8, 8);
        var shape = page.AddEllipse(4, 4, 3, .6, string.Empty, VisioMeasurementUnit.Inches);
        shape.FillColor = OfficeColor.Red; shape.LinePattern = 0; shape.Angle = degrees * Math.PI / 180;
        AssertRedOrientation(page, degrees);
    }

    [Theory]
    [InlineData(30)]
    [InlineData(-30)]
    public void StencilLineArtworkRotatesInThePageDirection(double degrees) {
        var page = VisioDocument.Create().AddPage("Artwork", 8, 8);
        var shape = page.AddRectangle(4, 4, 3, 3, string.Empty, VisioMeasurementUnit.Inches);
        shape.FillPattern = 0; shape.LinePattern = 0; shape.LineColor = OfficeColor.Red;
        shape.Angle = degrees * Math.PI / 180; shape.SetUserCell("OfficeIMO.StencilName", "Event", "STR");
        AssertRedOrientation(page, degrees);
    }

    [Theory]
    [InlineData(30)]
    [InlineData(-30)]
    public void PackagePreviewImageRotatesInThePageDirection(double degrees) {
        var page = VisioDocument.Create().AddPage("Preview", 8, 8);
        var master = new VisioMaster("preview", "Preview", new VisioShape("master", 0, 0, 3, .6, string.Empty));
        master.RawMasterRelationships.Add(new VisioAssets.MasterRelationshipContent {
            Id = "rIdImage", Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            Target = "../media/preview.custom", ContentType = "application/octet-stream", Extension = ".custom", Data = new byte[] { 1, 2, 3, 4 }
        });
        var shape = page.AddRectangle(4, 4, 3, .6, string.Empty, VisioMeasurementUnit.Inches);
        shape.Master = master; shape.NameU = master.NameU; shape.FillPattern = 0; shape.LinePattern = 0;
        shape.Angle = degrees * Math.PI / 180;
        shape.SetUserCell("OfficeIMO.StencilPreviewImageRelationshipId", "rIdImage", "STR");
        shape.SetUserCell("OfficeIMO.StencilPreviewImageTarget", "../media/preview.custom", "STR");
        AssertRedOrientation(page, degrees, new PreviewCodec());
    }

    private static void AssertRedOrientation(VisioPage page, double degrees, IOfficeRasterImageCodec? codec = null) {
        var image = VisualBaselineTestSupport.DecodePng(page.ToPng(new VisioPngSaveOptions {
            PixelsPerInch = 50, Supersampling = 1, ImageCodec = codec, ResolveConnectorLabelOverlaps = false
        }), "Rotated PNG must decode.");
        var points = new List<(double X, double Y)>();
        for (int y = 0; y < image.Height; y++)
            for (int x = 0; x < image.Width; x++) {
                var pixel = image.GetPixel(x, y);
                if (pixel.R > 180 && pixel.G < 170 && pixel.B < 170) points.Add((x, y));
            }
        Assert.True(points.Count > 30, "The rotated content must be visible.");
        double centerX = points.Average(point => point.X), centerY = points.Average(point => point.Y);
        double covariance = points.Sum(point => (point.X - centerX) * (point.Y - centerY));
        double varianceX = points.Sum(point => Math.Pow(point.X - centerX, 2));
        Assert.True(covariance * degrees < 0, $"Native rotation {degrees} degrees painted in the opposite screen direction; slope={covariance / varianceX}.");
    }

    private sealed class PreviewCodec : IOfficeRasterImageCodec {
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            image = new OfficeRasterImage(12, 2, OfficeColor.Red); return true;
        }
    }
}

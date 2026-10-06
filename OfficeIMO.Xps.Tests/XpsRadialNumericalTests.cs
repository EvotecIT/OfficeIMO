using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsRadialNumericalTests {
    [Fact]
    public void LargeRadiusEndpointPaintAndAlphaMaskKeepNonemptyPdfBounds() {
        var document = Create(XpsFormat.OpenXps, "Reflect", false);
        var page = document.Pages[0]; var markup = page.GetMarkup();
        foreach (var stop in markup.Descendants().Where(element => element.Name.LocalName == "GradientStop")) {
            string color = (string)stop.Attribute("Color")!;
            stop.SetAttributeValue("Color", "#80" + color.Substring(3));
        }
        page.ReplaceMarkup(markup);
        string syntax = Encoding.ASCII.GetString(document.ToPdf());
        var bounds = Regex.Matches(syntax, @"/BBox \[([^]]+)\]");
        Assert.Equal(2, bounds.Count); // RGB endpoint Form and gray luminosity mask.
        foreach (Match match in bounds) {
            var values = match.Groups[1].Value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
            Assert.True(values[2] > values[0]);
            Assert.True(values[3] > values[1]);
        }
        Assert.Contains("/SMask", syntax);
        Assert.DoesNotContain("/Subtype /Image", syntax);
    }

    [Theory]
    [InlineData(XpsFormat.Xps, "Pad", true)]
    [InlineData(XpsFormat.OpenXps, "Repeat", true)]
    [InlineData(XpsFormat.Xps, "Reflect", true)]
    [InlineData(XpsFormat.OpenXps, "Pad", false)]
    [InlineData(XpsFormat.Xps, "Repeat", false)]
    [InlineData(XpsFormat.OpenXps, "Reflect", false)]
    public void BoundedLargeRadiusFieldsRemainPaintedInRasterAndPdf(XpsFormat format, string spread, bool small) {
        var document = Create(format, spread, small);
        document = XpsDocument.Load(document.Save());
        var native = OfficeDrawingRasterRenderer.Render(document.Pages[0].ToDrawing(), background: OfficeColor.White);
        var pdfPage = Assert.Single(PdfReadDocument.Open(document.ToPdf()).Pages);
        var readback = OfficeDrawingRasterRenderer.Render(pdfPage.ToDrawing(), scale: 4D / 3D, background: OfficeColor.White);
        var point = small ? (102, 105) : (100, 80);
        foreach (var image in new[] { native, readback }) {
            var color = image.GetPixel(point.Item1, point.Item2);
            Assert.InRange((int)color.R, 251, 255);
            Assert.InRange((int)color.G, 0, 4);
            Assert.InRange((int)color.B, 0, 4);
            if (small) {
                var outside = image.GetPixel(105, 69);
                bool reflect = spread == "Reflect";
                Assert.InRange((int)(reflect ? outside.R : outside.B), 251, 255);
                Assert.InRange((int)(reflect ? outside.B : outside.R), 0, 4);
                Assert.InRange((int)outside.G, 0, 4);
            }
        }
    }

    internal static XpsDocument Create(XpsFormat format, string spread, bool small) {
        var document = XpsDocument.Create(format); var page = document.AddPage(200, 160);
        var xml = page.GetMarkup(); var ns = xml.Name.Namespace;
        var path = new XElement(ns + "Path", new XAttribute("Data", small ? "M100,79.5H100.1V80.5H100Z" : "M0,0H200V160H0Z"),
            new XAttribute("StrokeThickness", 0));
        if (small) path.SetAttributeValue("RenderTransform", "100,0,0,100,-9900,-7920");
        path.Add(new XElement(ns + "Path.Fill", new XElement(ns + "RadialGradientBrush",
            new XAttribute("MappingMode", "Absolute"),
            new XAttribute("Center", "-9999900,80"), new XAttribute("GradientOrigin", small ? "100.05,80" : "600,80"),
            new XAttribute("RadiusX", 10000000), new XAttribute("RadiusY", 10000000), new XAttribute("SpreadMethod", spread),
            new XElement(ns + "RadialGradientBrush.GradientStops",
                new XElement(ns + "GradientStop", new XAttribute("Offset", 0), new XAttribute("Color", "#FFFF0000")),
                new XElement(ns + "GradientStop", new XAttribute("Offset", 1), new XAttribute("Color", "#FF0000FF"))))));
        xml.Add(path); page.ReplaceMarkup(xml); return document;
    }
}

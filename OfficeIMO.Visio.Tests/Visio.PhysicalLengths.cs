using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio.Fluent;
using Xunit;

namespace OfficeIMO.Visio.Tests;

public sealed class PhysicalLengthApiTests {
    [Fact]
    public void MixedUnitsWriteNativeInchCells() {
        var document = VisioDocument.Create();
        var page = document.AddPage("Mixed", OfficeLength.Centimeters(21), OfficeLength.Millimeters(297));
        page.SetMargins(OfficeLength.Centimeters(1.27));
        var shape = page.AddRectangle(OfficeLength.Centimeters(2.54), OfficeLength.Points(144),
            OfficeLength.Millimeters(50.8), OfficeLength.Inches(1));
        Assert.Equal(OfficeLength.Inches(2), shape.Width);
        Assert.Equal(OfficeLength.Inches(2), shape.GetShapeBounds().Width);
        using var input = new MemoryStream(document.ToBytes());
        using var zip = new ZipArchive(input, ZipArchiveMode.Read);
        using var part = zip.GetEntry("visio/pages/page1.xml")!.Open();
        var xml = XDocument.Load(part);
        XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
        var nativeShape = xml.Descendants(ns + "Shape").Single();
        double Cell(XElement owner, string name) => double.Parse(owner.Elements(ns + "Cell")
            .Single(cell => (string?)cell.Attribute("N") == name).Attribute("V")!.Value, CultureInfo.InvariantCulture);
        Assert.Equal(1, Cell(nativeShape, "PinX"), 6);
        Assert.Equal(2, Cell(nativeShape, "PinY"), 6);
        Assert.Equal(2, Cell(nativeShape, "Width"), 6);
        Assert.Equal(1, Cell(nativeShape, "Height"), 6);
        var sheet = xml.Root!.Element(ns + "PageSheet")!;
        Assert.Equal(21 / 2.54, Cell(sheet, "PageWidth"), 6);
        Assert.Equal(29.7 / 2.54, Cell(sheet, "PageHeight"), 6);
        Assert.Equal(0.5, Cell(sheet, "PageLeftMargin"), 6);
    }

    [Fact]
    public void FluentCreationAndExplicitLookupUseTheObjectModelContract() {
        var document = VisioDocument.Create();
        var x = OfficeLength.Inches(1);
        var y = OfficeLength.Centimeters(5.08);
        var width = OfficeLength.Points(144);
        var height = OfficeLength.Millimeters(25.4);
        document.AsFluent().AddPage("Diagram", configure: page => page
            .AddRectangle("node", x, y, width, height, "Node")
            .SetMargins(OfficeLength.Inches(0.5)));
        var page = document.GetPage("diagram");
        var expected = new VisioPage("Reference").AddRectangle("node", x, y, width, height, "Node");
        var actual = Assert.Single(page.Shapes);
        Assert.Equal(expected.GetBounds(), actual.GetBounds());
        Assert.Equal(expected.NameU, actual.NameU);
        Assert.Equal(expected.Text, actual.Text);
        Assert.Same(page, document.GetOrAddPage("Diagram"));
        document.AsFluent().GetPage(0, existing => existing.GetShape("node", shape => shape.ApplyStyle(new VisioShapeStyle(OfficeColor.White, OfficeColor.Black))));
        Assert.Single(document.Pages);
        Assert.Same(page, document.GetPage(0));
        Assert.Throws<InvalidOperationException>(() => document.GetPage("missing"));
    }
}

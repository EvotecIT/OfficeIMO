using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgTests {
    [Fact]
    public void ReadsAndEditsLibreOfficeDrawingAndPreservesOpaquePreview() {
        string fixtures = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing");
        OdgDocument text = OdgDocument.LoadFlatXml(Path.Combine(fixtures, "libreoffice-transparent-text.fodg"));
        Assert.Equal("asdf", text.Pages[0].Shapes[0].Text);
        XElement geometry = new XElement(text.Pages[0].Shapes[0].ToXml().Element(OdfNamespaces.Draw + "enhanced-geometry")!);
        text.Pages[0].Shapes[0].Text = "Edited native drawing";
        OdgDocument reopened = OdgDocument.Load(new MemoryStream(text.ToBytes()));
        Assert.Equal("Edited native drawing", reopened.Pages[0].Shapes[0].Text);
        Assert.True(XNode.DeepEquals(geometry, reopened.Pages[0].Shapes[0].ToXml().Element(OdfNamespaces.Draw + "enhanced-geometry")));

        OdgDocument preview = OdgDocument.LoadFlatXml(Path.Combine(fixtures, "libreoffice-objectwithtext.fodg"));
        byte[] expected = preview.Pages[0].Shapes[0].GetImageBytes()!;
        Assert.Equal("VCLMTF", Encoding.ASCII.GetString(expected, 0, 6));
        using var flat = new MemoryStream(); preview.SaveFlatXml(flat); flat.Position = 0;
        Assert.Equal(expected, OdgDocument.LoadFlatXml(flat).Pages[0].Shapes[0].GetImageBytes());
        Assert.True(preview.Pages[0].ToDrawing().Report.HasSkippedOrUnsupported);
    }

    [Fact]
    public void EditsDrawingAndConvertsBetweenPackageAndFlatXml() {
        OdgDocument document = OdgDocument.Create();
        OdgPage page = document.AddPage("Overview", OdfLength.Centimeters(24), OdfLength.Centimeters(16));
        OdgShape box = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 5, 2), "Process");
        box.Text = "Receive  order\nCheck\tstock";
        box.FillColor = OdfColor.Parse("D1E9FF"); box.FontSize = OdfLength.Points(14);
        box.XmlId = "process";
        page.Shapes.AddLine(OdfLength.Centimeters(6), OdfLength.Centimeters(2), OdfLength.Centimeters(9), OdfLength.Centimeters(2));
        page.Shapes.AddGroup("Group").Children.AddEllipse(OdfRect.FromCentimeters(9, 1, 3, 2)).Text = "Ready";
        byte[] image = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        page.Shapes.AddImage(image, "pixel.png", OdfRect.FromCentimeters(1, 5, 1, 1));
        document.AddPage("Second"); document.MovePage(1, 0); document.RemovePage(0);
        Assert.True(document.Validate().IsValid);
        using var package = new MemoryStream(document.ToBytes());
        OdgDocument reopened = Assert.IsType<OdgDocument>(OdfDocument.Load(package));
        Assert.Equal("Receive  order\nCheck\tstock", reopened.Pages[0].Shapes[0].Text);
        Assert.Equal(box.FillColor, reopened.Pages[0].Shapes[0].FillColor);
        Assert.Equal("Ready", reopened.Pages[0].Shapes[2].Children[0].Text);
        Assert.Equal(image, reopened.Pages[0].Shapes[3].GetImageBytes());
        using var flat = new MemoryStream(); reopened.SaveFlatXml(flat); flat.Position = 0;
        OdgDocument flatDocument = OdgDocument.LoadFlatXml(flat);
        Assert.Equal(image, flatDocument.Pages[0].Shapes[3].GetImageBytes());
        flatDocument.Pages[0].Shapes[0].Text = "Edited";
        flatDocument.Pages[0].Width = OdfLength.Centimeters(25);
        OdgDocument edited = OdgDocument.Load(new MemoryStream(flatDocument.ToBytes()));
        Assert.Equal("Edited", edited.Pages[0].Shapes[0].Text);
        Assert.Equal(25, edited.Pages[0].Width.ToCentimeters(), 3);
        Assert.Throws<ArgumentException>(() => edited.Pages[0].Shapes[1].XmlId = "process");
        Assert.Throws<ArgumentException>(() => edited.AddPage("Overview"));
    }

    [Fact]
    public void PreservesUnsupportedDrawingXmlAndReportsProjectionLoss() {
        OdgDocument document = OdgDocument.Create();
        OdgPage page = document.AddPage();
        page.Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 4, 2), "Hello drawing");
        XElement custom = new XElement(OdfNamespaces.Draw + "custom-shape", new XAttribute(OdfNamespaces.Draw + "name", "Custom"),
            new XElement(OdfNamespaces.Draw + "enhanced-geometry", new XAttribute(OdfNamespaces.Draw + "type", "star5")));
        page.Element.Add(custom);
        page.Element.Add(new XElement(XName.Get("scene", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"), new XAttribute(OdfNamespaces.Draw + "name", "3D")));
        document.MarkPartDirty("content.xml");
        page.Shapes[0].Text = "Edited text";
        OdgDocument loaded = OdgDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.True(XNode.DeepEquals(custom, loaded.Pages[0].Shapes[1].ToXml()));
        OdfConversionResult<OfficeDrawing> projection = loaded.Pages[0].ToDrawing();
        Assert.True(projection.Report.HasSkippedOrUnsupported);
        Assert.Contains(projection.Report.Mappings, mapping => mapping.Feature == "shape:3D" && mapping.Status == OdfConversionMappingStatus.Skipped);
        Assert.Contains("Edited text", OfficeDrawingSvgExporter.ToSvg(projection.Value!, 1, OfficeSvgSizeUnit.Point));
        Assert.Throws<OdfConversionLossException>(() => loaded.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnAnyLoss));
    }

    [Fact]
    public void ReadsGraphicFontProperties() {
        OdgDocument document = OdgDocument.Create();
        OdgPage page = document.AddPage();
        OdgShape shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 4, 2));
        shape.Text = "Styled"; shape.FontSize = OdfLength.Points(18); shape.FontFamily = "Arial"; shape.TextColor = OdfColor.Parse("#123456");
        OdgShape read = OdgDocument.Load(new MemoryStream(document.ToBytes())).Pages[0].Shapes[0];
        Assert.Equal(18D, read.FontSize?.ToPoints());
        Assert.Equal("Arial", read.FontFamily); Assert.Equal(OdfColor.Parse("#123456"), read.TextColor);
    }

    [Fact]
    public void OrdersIndexedDrawingShapesAndKeepsGroupChildrenUnindexed() {
        OdgDocument document = OdgDocument.Create(); OdgPage page = document.AddPage();
        OdgShape front = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 4, 4), "Front"); front.Text = "Front";
        OdgShape back = page.Shapes.AddRectangle(OdfRect.FromCentimeters(2, 2, 4, 4), "Back"); back.Text = "Back";
        front.Element.SetAttributeValue(OdfNamespaces.Draw + "z-index", 5);
        back.Element.SetAttributeValue(OdfNamespaces.Draw + "z-index", 2);
        Assert.Equal(new[] { "Back", "Front" }, page.Shapes.Select(shape => shape.Name));
        Assert.Equal(new[] { "Back", "Front" }, page.ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>().Select(text => text.PlainText));
        page.Shapes.Move(0, 1);
        OdgShape added = page.Shapes.AddEllipse(OdfRect.FromCentimeters(3, 3, 1, 1), "Last");
        Assert.Equal(new[] { "Front", "Back", "Last" }, page.Shapes.Select(shape => shape.Name));
        var group = page.Shapes.AddGroup(); group.Children.AddRectangle(OdfRect.FromCentimeters(1, 1, 1, 1)); group.Children.AddEllipse(OdfRect.FromCentimeters(2, 2, 1, 1)); group.Children.Move(0, 1);
        Assert.All(group.Children, child => Assert.Null(child.ToXml().Attribute(OdfNamespaces.Draw + "z-index")));
        var reopened = OdgDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(new[] { "Front", "Back", "Last" }, reopened.Pages[0].Shapes.Take(3).Select(shape => shape.Name));
        XElement[] xmlOrder = reopened.Pages[0].Element.Elements().ToArray();
        Assert.True(int.Parse((string)xmlOrder[0].Attribute(OdfNamespaces.Draw + "z-index")!) < int.Parse((string)xmlOrder[1].Attribute(OdfNamespaces.Draw + "z-index")!));
    }

    [Fact]
    public void FlatDrawingRejectsEntitiesWrongFamilyAndResourceLimitViolations() {
        using var entities = new MemoryStream(Encoding.UTF8.GetBytes("<!DOCTYPE x [<!ENTITY value 'expanded'>]><x>&value;</x>"));
        Assert.ThrowsAny<Exception>(() => OdgDocument.LoadFlatXml(entities));
        OdgDocument drawing = OdgDocument.Create(); drawing.AddPage();
        using var smallLimit = new MemoryStream(drawing.ToBytes());
        Assert.ThrowsAny<Exception>(() => OdgDocument.Load(smallLimit, new OdfLoadOptions { MaxPackageBytes = 10 }));
        using var wrongFamily = new MemoryStream(OdpPresentation.Create().ToBytes());
        Assert.Throws<InvalidDataException>(() => OdgDocument.Load(wrongFamily));
    }
}

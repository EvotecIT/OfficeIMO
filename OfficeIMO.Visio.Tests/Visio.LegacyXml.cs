using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioLegacyXmlTests {
    [Fact]
    public void ReadsIndependentLegacyDrawingAndPreservesUntouchedSourceShapeContent() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "clickhouse-replication.vdx");
        var imported = VisioDocument.LoadLegacyXml(path);
        VisioPage page = Assert.Single(imported.Value.Pages);
        Assert.Equal(19, page.Shapes.Count + page.Connectors.Count);
        Assert.DoesNotContain(imported.Report.FidelityDiagnostics, item => item.Code == "VDX_CONNECTOR_PROFILE" && item.Location == "34");
        VisioConnector connected = Assert.Single(page.Connectors, connector => connector.Id == "28");
        Assert.Equal("4", connected.From?.Id); Assert.Equal("11", connected.To?.Id);
        Assert.Contains(page.Shapes, shape => !string.IsNullOrWhiteSpace(shape.Text));
        string?[] untouched = page.Shapes.Skip(1).Select(shape => shape.Text).ToArray();
        page.Shapes[0].Text = "OfficeIMO edit";
        var exported = imported.Value.ToLegacyXmlResult();
        VisioDocument reopened = VisioDocument.LoadLegacyXml(new MemoryStream(exported.Value)).Value;
        Assert.Equal("OfficeIMO edit", reopened.Pages[0].Shapes[0].Text);
        Assert.Equal(19, reopened.Pages[0].Shapes.Count + reopened.Pages[0].Connectors.Count);
        Assert.DoesNotContain("data-officeimo-stencil-artwork", reopened.Pages[0].ToSvg());
        Assert.Null(reopened.FilePath);
        Assert.Equal(untouched, reopened.Pages[0].Shapes.Skip(1).Select(shape => shape.Text));
        XNamespace ns = "http://schemas.microsoft.com/visio/2003/core";
        XDocument xml = XDocument.Load(new MemoryStream(exported.Value), LoadOptions.PreserveWhitespace);
        XElement pageXml = Assert.Single(xml.Descendants(ns + "Page"));
        Assert.Equal(19, pageXml.Element(ns + "Shapes")!.Elements(ns + "Shape").Count());
        Assert.Contains(pageXml.Element(ns + "Shapes")!.Elements(ns + "Shape"), shape => (string?)shape.Attribute("ID") == "34");
        Assert.All(xml.Descendants(ns + "Shape"), shape => {
            var children = shape.Elements().ToArray();
            int nested = Array.FindIndex(children, child => child.Name == ns + "Shapes");
            Assert.True(nested < 0 || nested == children.Length - 1);
        });
    }

    [Fact]
    public void ReportsOmissionsAndLeavesDestinationUntouchedUntilExplicitlyAccepted() {
        VisioDocument document = VisioDocument.Create(); document.AddPage("Page", 10, 7);
        using var target = new MemoryStream(new byte[100], writable: true);
        byte[] before = target.ToArray();
        Assert.Throws<OfficeConversionException>(() => document.SaveLegacyXml(target));
        Assert.Equal(before, target.ToArray());
    }

    [Theory]
    [InlineData(VisioPackageType.Drawing)]
    [InlineData(VisioPackageType.Template)]
    [InlineData(VisioPackageType.Stencil)]
    public void CreatesEditsAndConvertsLegacyFamilies(VisioPackageType family) {
        VisioDocument document = VisioDocument.Create(family);
        if (family == VisioPackageType.Stencil) document.RegisterMaster("Box", new VisioShape("1", 2, 2, 2, 1, "Reusable"));
        else {
            VisioPage page = document.AddPage("Workflow", 10, 7);
            var start = new VisioShape("1", 2, 2, 2, 1, "Start");
            var end = new VisioShape("2", 6, 2, 2, 1, "End");
            page.Shapes.Add(start); page.Shapes.Add(end);
            page.Connectors.Add(new VisioConnector(start, end));
        }
        var exported = document.ToLegacyXmlResult();
        Assert.DoesNotContain(exported.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "VDX_GEOMETRY" || diagnostic.Code == "VDX_SECTION");
        XDocument xml = XDocument.Load(new MemoryStream(exported.Value));
        Assert.Equal("http://schemas.microsoft.com/visio/2003/core", xml.Root!.Name.NamespaceName);
        Assert.DoesNotContain(xml.Descendants(), element => element.Name.LocalName == "Cell");
        Assert.All(xml.Root.Elements(), element => Assert.Null(element.Attribute(XNamespace.Xml + "space")));
        XNamespace ns = xml.Root.Name.Namespace;
        Assert.All(xml.Descendants(ns + "Geom").Elements().Where(e => e.Name.LocalName == "MoveTo" || e.Name.LocalName == "LineTo"),
            row => Assert.NotNull(row.Attribute("IX")));
        Assert.All(xml.Descendants(ns + "Master"), master => Assert.Null(master.Attribute("MasterType")));
        var imported = VisioDocument.LoadLegacyXml(new MemoryStream(exported.Value), family);
        Assert.Equal(family, imported.Value.PackageType);
        if (family == VisioPackageType.Stencil) Assert.Equal("Reusable", Assert.Single(imported.Value.Masters).Shape.Text);
        else {
            VisioPage page = Assert.Single(imported.Value.Pages);
            Assert.Equal(2, page.Shapes.Count); Assert.Single(page.Connectors);
            page.Shapes[0].Text = "Edited";
            using var legacy = new MemoryStream(); imported.Value.SaveLegacyXml(legacy, allowOmissions: true); legacy.Position = 0;
            VisioDocument edited = VisioDocument.LoadLegacyXml(legacy, family).Value;
            Assert.Equal("Edited", edited.Pages[0].Shapes[0].Text);
            VisioDocument package = VisioDocument.Load(new MemoryStream(edited.ToBytes()));
            Assert.Equal("Edited", package.Pages[0].Shapes[0].Text);
        }
    }

    [Fact]
    public void PreservesNamedRowsAndResolvesLegacyPaletteWithoutChangingRichTextSpacing() {
        const string source = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Colors><ColorEntry IX='7' RGB='#12AB34'/></Colors><Pages><Page ID='0' Name='Page'><PageSheet><Layer IX='0'><Name>Review</Name><Color>7</Color></Layer></PageSheet><Shapes><Shape ID='1' Type='Shape'><XForm><PinX>1</PinX><PinY>1</PinY><Width>2</Width><Height>1</Height></XForm><Fill><FillForegnd>7</FillForegnd></Fill><User NameU='Example' ID='3'><Value F='42'>42</Value></User><Hyperlink NameU='Link' ID='5'><Address>https://example.org</Address></Hyperlink><Text><cp IX='0'/>first<cp IX='1'/> second</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(source));
        stream.Position = 9;
        var imported = VisioDocument.LoadLegacyXml(stream);
        Assert.Equal(9, stream.Position);
        Assert.True(stream.CanRead);
        Assert.Equal(7, imported.Value.Pages[0].Layers[0].Color);
        var output = imported.Value.ToLegacyXmlResult();
        XNamespace ns = "http://schemas.microsoft.com/visio/2003/core";
        XDocument xml = XDocument.Load(new MemoryStream(output.Value), LoadOptions.PreserveWhitespace);
        Assert.Equal("#12AB34", xml.Descendants(ns + "FillForegnd").Last().Value);
        Assert.Equal("7", Assert.Single(xml.Descendants(ns + "Layer")).Element(ns + "Color")!.Value);
        XElement row = Assert.Single(xml.Descendants(ns + "User"));
        Assert.Equal("3", (string?)row.Attribute("ID"));
        Assert.Equal("Example", (string?)row.Attribute("NameU"));
        Assert.Equal("5", (string?)Assert.Single(xml.Descendants(ns + "Hyperlink")).Attribute("ID"));
        Assert.Equal("first second", Assert.Single(xml.Descendants(ns + "Text")).Value);
        using var cancelled = new System.Threading.CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => VisioDocument.LoadLegacyXml(stream, cancellationToken: cancelled.Token));
        Assert.Equal(9, stream.Position);
    }

    [Fact]
    public void RejectsEntitiesExcessDepthAndWrongNamespaceBeforeImport() {
        using var entity = new MemoryStream(Encoding.UTF8.GetBytes("<!DOCTYPE x [<!ENTITY a 'b'>]><x>&a;</x>"));
        Assert.ThrowsAny<Exception>(() => VisioDocument.LoadLegacyXml(entity));
        using var wrong = new MemoryStream(Encoding.UTF8.GetBytes("<VisioDocument/>"));
        Assert.Throws<InvalidDataException>(() => VisioDocument.LoadLegacyXml(wrong));
        using var deep = new MemoryStream(Encoding.UTF8.GetBytes("<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Pages><Page><Shapes/></Page></Pages></VisioDocument>"));
        Assert.Throws<InvalidDataException>(() => VisioDocument.LoadLegacyXml(deep, options: new VisioLoadOptions { MaxLegacyXmlDepth = 2 }));
    }
}

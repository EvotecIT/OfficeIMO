using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioTextStylePreservationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InheritedFormattingDoesNotBecomeAnOverrideUntilEdited(bool localRows) {
        XNamespace legacy = "http://schemas.microsoft.com/visio/2003/core";
        const string master = "<Shape ID='1'><XForm><Width>2</Width><Height>1</Height></XForm><Char IX='7'><Color>#000000</Color><Size Unit='PT' F='GUARD(12 pt)'>0.16666666666666667</Size></Char><Para IX='4'><HorzAlign>0</HorzAlign></Para><Text><cp IX='7'/><pp IX='4'/>Inherited</Text></Shape>";
        string local = localRows ? "<Char IX='7'><Size Unit='PT' F='Inh'>0.16666666666666667</Size></Char><Para IX='4'><HorzAlign F='Inh'>0</HorzAlign></Para>" : "";
        string source = $"<VisioDocument xmlns='{legacy}'><Masters><Master ID='0' NameU='Styled'><Shapes>{master}</Shapes></Master></Masters><Pages><Page ID='0'><Shapes><Shape ID='8' Master='0'><XForm><PinX>3</PinX><PinY>2</PinY></XForm>{local}</Shape></Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
        for (int round = 0; round < 3; round++) {
            var pageShape = ExportShape(document);
            Assert.Equal(localRows ? 1 : 0, pageShape.Elements(legacy + "Char").Count());
            Assert.Equal(localRows ? 1 : 0, pageShape.Elements(legacy + "Para").Count());
            Assert.Null(pageShape.Element(legacy + "Text"));
            if (localRows) {
                Assert.Single(pageShape.Element(legacy + "Char")!.Elements());
                Assert.Equal("Inh", (string?)pageShape.Element(legacy + "Char")!.Element(legacy + "Size")!.Attribute("F"));
                Assert.Equal("Inh", (string?)pageShape.Element(legacy + "Para")!.Element(legacy + "HorzAlign")!.Attribute("F"));
            }
            document = VisioDocument.Load(new MemoryStream(document.ToBytes()));
        }
        var style = document.Pages[0].Shapes[0].TextStyle!;
        style.Color = OfficeColor.Red;
        style.HorizontalAlignment = VisioTextHorizontalAlignment.Right;
        foreach (var candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            var shape = ExportShape(candidate);
            var character = Assert.Single(shape.Elements(legacy + "Char"));
            Assert.Equal("7", (string?)character.Attribute("IX"));
            Assert.Equal("#FF0000", character.Element(legacy + "Color")!.Value);
            if (localRows) Assert.Equal("Inh", (string?)character.Element(legacy + "Size")!.Attribute("F"));
            else Assert.Null(character.Element(legacy + "Size"));
            var paragraph = Assert.Single(shape.Elements(legacy + "Para"));
            Assert.Equal("4", (string?)paragraph.Attribute("IX"));
            Assert.Equal("2", paragraph.Element(legacy + "HorzAlign")!.Value);
            Assert.Null(paragraph.Element(legacy + "HorzAlign")!.Attribute("F"));
            Assert.Null(shape.Element(legacy + "Text"));
        }

        document.Pages[0].Shapes[0].Text = "Edited label";
        Assert.Equal("Edited label", ExportShape(VisioDocument.Load(new MemoryStream(document.ToBytes()))).Element(legacy + "Text")!.Value);
        document.Pages[0].Shapes[0].Text = "";
        Assert.Equal("", ExportShape(VisioDocument.Load(new MemoryStream(document.ToBytes()))).Element(legacy + "Text")!.Value);

        XElement ExportShape(VisioDocument candidate) => XDocument.Load(new MemoryStream(candidate.ToLegacyXmlResult().Value))
            .Descendants(legacy + "Pages").Descendants(legacy + "Shape").Single();
    }

    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ModeledStyleEditsKeepNativeRowIndicesTextMarkersAndUntouchedFormulas(bool connector, bool package) {
        var document = Load(connector, charIndex: 7, paraIndex: 4);
        VisioTextStyle style = connector ? document.Pages[0].Connectors[0].TextStyle! : document.Pages[0].Shapes[0].TextStyle!;
        style.Color = OfficeColor.Red;
        style.HorizontalAlignment = VisioTextHorizontalAlignment.Right;
        if (package) document = VisioDocument.Load(new MemoryStream(document.ToBytes()));
        XElement shape = SavedShape(document);
        XElement character = Assert.Single(shape.Elements(Legacy + "Char"));
        XElement paragraph = Assert.Single(shape.Elements(Legacy + "Para"));
        Assert.Equal("7", (string?)character.Attribute("IX"));
        Assert.Equal("4", (string?)paragraph.Attribute("IX"));
        Assert.Equal("#FF0000", character.Element(Legacy + "Color")!.Value);
        Assert.Null(character.Element(Legacy + "Color")!.Attribute("F"));
        Assert.Equal("GUARD(12 pt)", (string?)character.Element(Legacy + "Size")!.Attribute("F"));
        Assert.Equal("Inh", (string?)character.Element(Legacy + "Style")!.Attribute("F"));
        Assert.Equal("2", paragraph.Element(Legacy + "HorzAlign")!.Value);
        Assert.Null(paragraph.Element(Legacy + "HorzAlign")!.Attribute("F"));
        Assert.Equal("7", (string?)shape.Element(Legacy + "Text")!.Element(Legacy + "cp")!.Attribute("IX"));
        Assert.Equal("4", (string?)shape.Element(Legacy + "Text")!.Element(Legacy + "pp")!.Attribute("IX"));
        Assert.Equal("Label", shape.Element(Legacy + "Text")!.Value);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void UneditedModeledTextRowsRetainTheirNativeCellsAndFormulas(bool connector, bool package) {
        var document = Load(connector, charIndex: 0, paraIndex: 0);
        if (package) document = VisioDocument.Load(new MemoryStream(document.ToBytes()));
        XElement shape = SavedShape(document);
        XElement character = Assert.Single(shape.Elements(Legacy + "Char"));
        XElement paragraph = Assert.Single(shape.Elements(Legacy + "Para"));
        Assert.Equal("RGB(0,0,0)", (string?)character.Element(Legacy + "Color")!.Attribute("F"));
        Assert.Equal("GUARD(12 pt)", (string?)character.Element(Legacy + "Size")!.Attribute("F"));
        Assert.Equal("Inh", (string?)character.Element(Legacy + "Style")!.Attribute("F"));
        Assert.Equal("0", (string?)paragraph.Element(Legacy + "HorzAlign")!.Attribute("F"));
        Assert.Equal(new[] { "Color", "Style", "Size" }, character.Elements().Select(e => e.Name.LocalName));
    }

    [Theory]
    [InlineData(VisioPackageType.Stencil, false)]
    [InlineData(VisioPackageType.Stencil, true)]
    [InlineData(VisioPackageType.Template, false)]
    [InlineData(VisioPackageType.Template, true)]
    public void LoadedMasterStyleEditsRetainNativeRowIdentity(VisioPackageType family, bool package) {
        var document = Load(false, 7, 4, master: true, family: family);
        document.Masters.Single().Shape.TextStyle!.Color = OfficeColor.Red;
        if (package) document = VisioDocument.Load(new MemoryStream(document.ToBytes()));
        XElement shape = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value))
            .Descendants(Legacy + "Masters").Descendants(Legacy + "Shape").Single();
        XElement character = Assert.Single(shape.Elements(Legacy + "Char"));
        Assert.Equal("7", (string?)character.Attribute("IX"));
        Assert.Equal("#FF0000", character.Element(Legacy + "Color")!.Value);
        Assert.Null(character.Element(Legacy + "Color")!.Attribute("F"));
        Assert.Equal("GUARD(12 pt)", (string?)character.Element(Legacy + "Size")!.Attribute("F"));
        Assert.Equal("4", (string?)Assert.Single(shape.Elements(Legacy + "Para")).Attribute("IX"));
    }

    [Fact]
    public void ClearingModeledFormattingRetainsRowsReferencedByTextMarkers() {
        var document = Load(false, 7, 4);
        document.Pages[0].Shapes[0].TextStyle = new VisioTextStyle();
        XElement shape = SavedShape(document);
        XElement character = Assert.Single(shape.Elements(Legacy + "Char"));
        XElement paragraph = Assert.Single(shape.Elements(Legacy + "Para"));
        Assert.Equal("7", (string?)character.Attribute("IX"));
        Assert.Equal("4", (string?)paragraph.Attribute("IX"));
        Assert.Empty(character.Elements());
        Assert.Empty(paragraph.Elements());
        Assert.Equal("Label", shape.Element(Legacy + "Text")!.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MasterInstanceStyleChangesKeepRemappedFormulasAndDoNotEditTheSource(bool group) {
        var document = Load(false, 7, 4, master: true, group: group, sizeFormula: "GUARD(Sheet.1!Width*6 pt)");
        var master = document.Masters.Single();
        var instance = document.AddPage("Instance").AddShape("8", master, 3, 3, 2, 1);
        instance.TextStyle!.Color = OfficeColor.Red;
        var xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        XElement shape = xml.Descendants(Legacy + "Pages").Descendants(Legacy + "Shape").Single(e => (string?)e.Attribute("ID") == "8");
        XElement character = Assert.Single(shape.Elements(Legacy + "Char"));
        Assert.Equal("7", (string?)character.Attribute("IX"));
        Assert.Equal("#FF0000", character.Element(Legacy + "Color")!.Value);
        Assert.Equal("GUARD(Sheet.8!Width*6 pt)", (string?)character.Element(Legacy + "Size")!.Attribute("F"));
        XElement original = xml.Descendants(Legacy + "Masters").Descendants(Legacy + "Shape").First();
        Assert.Equal("#000000", original.Element(Legacy + "Char")!.Element(Legacy + "Color")!.Value);
        Assert.Equal("GUARD(Sheet.1!Width*6 pt)", (string?)original.Element(Legacy + "Char")!.Element(Legacy + "Size")!.Attribute("F"));
    }

    private static VisioDocument Load(bool connector, int charIndex, int paraIndex, bool master = false,
        VisioPackageType family = VisioPackageType.Stencil, bool group = false, string sizeFormula = "GUARD(12 pt)") {
        string oneD = connector ? "<XForm1D><BeginX>1</BeginX><BeginY>1</BeginY><EndX>3</EndX><EndY>1</EndY></XForm1D>" : "";
        string source = $"<VisioDocument xmlns='{Legacy}'><Pages><Page ID='0'><Shapes><Shape ID='1' Type='Shape'><XForm><PinX>2</PinX><PinY>1</PinY><Width>2</Width><Height>1</Height></XForm>{oneD}<Char IX='{charIndex}'><Color F='RGB(0,0,0)'>#000000</Color><Style F='Inh'>0</Style><Size Unit='PT' F='{sizeFormula}'>0.16666666666666667</Size></Char><Para IX='{paraIndex}'><HorzAlign F='0'>0</HorzAlign></Para><Text><cp IX='{charIndex}'/><pp IX='{paraIndex}'/>Label</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        if (master) {
            var xml = XDocument.Parse(source);
            XElement shape = xml.Descendants(Legacy + "Shape").Single();
            shape.Remove();
            if (group) {
                shape.SetAttributeValue("Type", "Group");
                shape.Add(new XElement(Legacy + "Shapes", new XElement(Legacy + "Shape", new XAttribute("ID", "2"),
                    new XElement(Legacy + "Text", "Child"))));
            }
            xml.Root!.ReplaceNodes(new XElement(Legacy + "Masters", new XElement(Legacy + "Master",
                new XAttribute("ID", "1"), new XAttribute("NameU", "Label"), new XElement(Legacy + "Shapes", shape))));
            source = xml.ToString();
        }
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source)),
            packageType: master ? family : VisioPackageType.Drawing).Value;
    }

    private static XElement SavedShape(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value))
        .Descendants(Legacy + "Pages").Descendants(Legacy + "Shape").Single();
}

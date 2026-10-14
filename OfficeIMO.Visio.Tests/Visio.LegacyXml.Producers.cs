using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioLegacyProducerTests {
    private const string Namespace2002 = "urn:schemas-microsoft-com:office:visio";
    private const string Namespace2003 = "http://schemas.microsoft.com/visio/2003/core";

    [Theory]
    [InlineData("nxbre-ie3.vsx", VisioPackageType.Stencil, 12)]
    [InlineData("pronom-visio2002.vsx", VisioPackageType.Stencil, 1)]
    [InlineData("pronom-visio2002.vtx", VisioPackageType.Template, 1)]
    public void IndependentStencilsAndTemplatesKeepMastersAndNamedRowsAfterEdit(string file, VisioPackageType family, int masterCount) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", file);
        XDocument original = XDocument.Load(path, LoadOptions.PreserveWhitespace);
        var result = VisioDocument.LoadLegacyXml(path);
        var document = result.Value;
        Assert.Equal(family, document.PackageType); Assert.Equal(masterCount, document.Masters.Count);
        if (original.Root!.Name.NamespaceName == Namespace2002)
            Assert.Contains(result.Report.FidelityDiagnostics, d => d.Code == "VDX_2002_INPUT");
        var unedited = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        Assert.Equal(MasterPageSheets(original), MasterPageSheets(unedited));
        Assert.Equal(original.Descendants(original.Root.Name.Namespace + "Masters").Descendants(original.Root.Name.Namespace + "Text").Select(e => e.Value),
            unedited.Descendants(XName.Get("Masters", Namespace2003)).Descendants(XName.Get("Text", Namespace2003)).Select(e => e.Value));
        var master = document.Masters.First();
        int nested = master.Shape.Children.Count;
        master.Shape.Text = "Edited master label";
        master.Shape.SetUserCell("OfficeIMOReview", "accepted", "STR");
        foreach (var candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            var saved = candidate.ToLegacyXmlResult();
            var xml = XDocument.Load(new MemoryStream(saved.Value), LoadOptions.PreserveWhitespace);
            Assert.Equal(Namespace2003, xml.Root!.Name.NamespaceName);
            Assert.Equal(MasterPageSheets(original), MasterPageSheets(xml));
            XNamespace sourceNs = original.Root.Name.Namespace;
            foreach (XElement sourceMaster in original.Descendants(sourceNs + "Masters").Elements(sourceNs + "Master")) {
                XElement savedMaster = xml.Descendants(XName.Get("Master", Namespace2003)).Single(m => (string?)m.Attribute("NameU") == (string?)sourceMaster.Attribute("NameU"));
                foreach (XAttribute attribute in sourceMaster.Attributes().Where(a => !a.IsNamespaceDeclaration && a.Name != "ID"))
                    Assert.Equal(attribute.Value, (string?)savedMaster.Attribute(attribute.Name));
            }
            Assert.Equal(NamedRows(original), NamedRows(xml).Where(row => !row.Contains("OfficeIMOReview")).ToArray());
            var reopened = VisioDocument.LoadLegacyXml(new MemoryStream(saved.Value), family).Value;
            Assert.Equal(masterCount, reopened.Masters.Count);
            Assert.Equal("Edited master label", reopened.Masters.First().Shape.Text);
            Assert.Equal("accepted", reopened.Masters.First().Shape.GetUserCellValue("OfficeIMOReview"));
            Assert.Equal(nested, reopened.Masters.First().Shape.Children.Count);
            Assert.Equal(document.Pages.Count, reopened.Pages.Count);
        }
    }

    [Theory]
    [InlineData(Namespace2002)]
    [InlineData(Namespace2003)]
    public void LegacyFontTableDrivesTextAndRetainsNativeMetadata(string ns) {
        string source = "<VisioDocument xmlns='" + ns + "'><Fonts><FontEntry ID='3' Name='Consolas' CharSet='0' PitchAndFamily='49' Attributes='123' Weight='400' Unicode='1'/></Fonts><Pages><Page ID='0'><Shapes><Shape ID='1'><XForm><PinX>1</PinX><PinY>1</PinY><Width>2</Width><Height>1</Height></XForm><Char IX='0'><Font>3</Font><Size Unit='PT'>0.1666666666666667</Size></Char><Text>original</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        var document = Load(source);
        var shape = document.Pages[0].Shapes[0];
        Assert.Equal("Consolas", shape.TextStyle!.FontFamily);
        shape.Text = "edited";
        var reopened = VisioDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal("Consolas", reopened.Pages[0].Shapes[0].TextStyle!.FontFamily);
        var xml = XDocument.Load(new MemoryStream(reopened.ToLegacyXmlResult().Value));
        var font = Assert.Single(xml.Descendants(XName.Get("FontEntry", Namespace2003)));
        Assert.Equal("49", (string?)font.Attribute("PitchAndFamily"));
        Assert.Equal("123", (string?)font.Attribute("Attributes"));
        Assert.Equal("Consolas", (string?)font.Attribute("Name"));
    }

    [Fact]
    public void OlderNamespaceKeepsMacroAndExternalXmlRejection() {
        Assert.Throws<NotSupportedException>(() => Load("<VisioDocument xmlns='" + Namespace2002 + "'><VBProjectData>AAAA</VBProjectData></VisioDocument>"));
        Assert.ThrowsAny<Exception>(() => Load("<!DOCTYPE x [<!ENTITY x SYSTEM 'file:///not-accessed'>]><VisioDocument xmlns='" + Namespace2002 + "'>&x;</VisioDocument>"));
        Assert.Throws<InvalidDataException>(() => Load("<VisioDocument xmlns='urn:unrecognized:visio'/>"));
    }

    [Fact]
    public void RegisteredMasterCreationUsesPageUnitWithoutAmbiguousOverloads() {
        var document = VisioDocument.Create(); var page = document.AddPage("Instances");
        page.DefaultUnit = VisioMeasurementUnit.Centimeters;
        document.RegisterMaster("Box", new VisioShape("1", 0, 0, 1, 1, ""));
        var plain = page.AddShape("1", "Box", 2.54, 5.08, 7.62, 2.54);
        var text = page.AddShape("2", "Box", 2.54, 5.08, 7.62, 2.54, "label");
        var explicitUnit = page.AddShape("3", "Box", 1, 2, 3, 1, unit: VisioMeasurementUnit.Inches);
        Assert.Equal(1, plain.PinX, 8); Assert.Equal(2, plain.PinY, 8); Assert.Equal(3, plain.Width, 8);
        Assert.Equal(plain.Width, text.Width); Assert.Equal(plain.Width, explicitUnit.Width);
        Assert.Equal("label", text.Text);
    }

    [Fact]
    public void LoadedMasterEditsKeepUntouchedFormulasAndSaveNestedChangesAndNewFont() {
        string source = "<VisioDocument xmlns='" + Namespace2003 + "'><Fonts><FontEntry ID='0' Name='Arial' CharSet='0'/><FontEntry ID='3' Name='Arial' CharSet='2'/></Fonts><Masters><Master ID='1' NameU='Group'><Shapes><Shape ID='1' Type='Group'><XForm><Width F='GUARD(2)'>2</Width><Height F='GUARD(1)'>1</Height></XForm><Shapes><Shape ID='2'><Text>removed</Text></Shape><Shape ID='3'><XForm><Width>1</Width><Height>1</Height></XForm><Char IX='0'><Font>3</Font></Char><Text>original</Text></Shape></Shapes></Shape></Shapes></Master><Master ID='2' NameU='Unused'><Shapes><Shape ID='1'><Text>retained</Text></Shape></Shapes></Master></Masters></VisioDocument>";
        var document = Load(source, VisioPackageType.Template);
        var root = document.Masters.First().Shape;
        root.Width = 4;
        root.Children.RemoveAt(0);
        root.Children[0].Text = "edited child";
        root.Children[0].TextStyle!.FontFamily = "Consolas";
        root.Children.Insert(0, new VisioShape("4", 1, 1, 1, 1, "new child"));
        var saved = document.ToLegacyXmlResult();
        XNamespace ns = Namespace2003;
        var xml = XDocument.Load(new MemoryStream(saved.Value));
        var transform = xml.Descendants(ns + "Shape").First().Element(ns + "XForm")!;
        Assert.Equal("4", transform.Element(ns + "Width")!.Value);
        Assert.Null(transform.Element(ns + "Width")!.Attribute("F"));
        Assert.Equal("GUARD(1)", (string?)transform.Element(ns + "Height")!.Attribute("F"));
        var reopened = VisioDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(2, reopened.Masters.Count); // Templates retain unused masters for later authoring.
        var editedRoot = reopened.Masters.First().Shape;
        Assert.Equal(new[] { "4", "3" }, editedRoot.Children.Select(child => child.Id));
        Assert.Equal("edited child", editedRoot.Children[1].Text);
        Assert.Equal("Consolas", editedRoot.Children[1].TextStyle!.FontFamily);
        // Repeated save uses the same source baseline without reintroducing removed children.
        var twice = VisioDocument.LoadLegacyXml(new MemoryStream(reopened.ToLegacyXmlResult().Value), VisioPackageType.Template).Value;
        Assert.Equal(new[] { "new child", "edited child" }, twice.Masters.First().Shape.Children.Select(child => child.Text));
    }

    private static VisioDocument Load(string source, VisioPackageType family = VisioPackageType.Drawing) => VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source)), family).Value;
    private static string[] MasterPageSheets(XDocument document) {
        XNamespace ns = document.Root!.Name.Namespace;
        return document.Descendants(ns + "Masters").Elements(ns + "Master")
            .Select(master => (string?)master.Attribute("NameU") + "/" + Canonical(master.Element(ns + "PageSheet")!))
            .OrderBy(value => value, StringComparer.Ordinal).ToArray();
    }
    private static string[] NamedRows(XDocument document) {
        XNamespace ns = document.Root!.Name.Namespace;
        return document.Descendants(ns + "Masters").Elements(ns + "Master").SelectMany(master =>
            master.Descendants(ns + "Shape").SelectMany(shape => shape.Elements().Where(row => row.Name == ns + "User" || row.Name == ns + "Prop" || row.Name == ns + "Hyperlink")
                .Select(row => (string?)master.Attribute("NameU") + "/" + (string?)shape.Attribute("ID") + "/" + Canonical(row))))
            .OrderBy(value => value, StringComparer.Ordinal).ToArray();
    }
    private static string Canonical(XElement source) {
        XElement copy = new XElement(source);
        foreach (var element in copy.DescendantsAndSelf()) {
            element.Name = element.Name.LocalName;
            if (element.HasElements) element.Nodes().OfType<XText>().Where(text => string.IsNullOrWhiteSpace(text.Value)).Remove();
            // Empty element syntax and an explicit empty text value have the same ShapeSheet meaning.
            if (!element.HasElements && element.Value.Length == 0) element.RemoveNodes();
            var attributes = element.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString(), StringComparer.Ordinal).Select(a => new XAttribute(a)).ToArray();
            element.ReplaceAttributes(attributes);
        }
        return copy.ToString(SaveOptions.DisableFormatting);
    }
}

using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioShapeSheetSectionPreservationTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    public void LoadedSectionsSupportReplacementRemovalAndAddition(bool connector, bool package, bool directCellEdit) {
        string oneD = connector ? "<XForm1D><BeginX>1</BeginX><BeginY>1</BeginY><EndX>3</EndX><EndY>1</EndY></XForm1D>" : "";
        string source = $"<VisioDocument xmlns='{Legacy}'><Pages><Page ID='0'><Shapes><Shape ID='1' Type='Shape'><XForm><PinX>2</PinX><PinY>1</PinY><Width>2</Width><Height>1</Height></XForm>{oneD}<Scratch IX='0'><X F='Sheet.99!Width' Err='#REF!'>1</X><Y F='Inh'>2</Y></Scratch><Text>Label</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
        var shape = connector ? null : document.Pages[0].Shapes[0];
        var edge = connector ? document.Pages[0].Connectors[0] : null;
        var section = (shape?.GetShapeSheetSections() ?? edge!.GetShapeSheetSections()).Single(s => s.Name == "Scratch");
        if (directCellEdit) {
            var cell = section.Rows.Single().FindCell("X")!;
            cell.Value = "3";
            cell.Formula = null;
        } else {
            section.Rows.Single().SetCell("X", "3");
        }
        Set(section);
        XElement edited = XmlShape(package);
        XElement scratch = Assert.Single(edited.Elements(Legacy + "Scratch"));
        Assert.Equal("3", scratch.Element(Legacy + "X")!.Value);
        Assert.Null(scratch.Element(Legacy + "X")!.Attribute("F"));
        Assert.Null(scratch.Element(Legacy + "X")!.Attribute("Err"));
        Assert.Equal("Inh", (string?)scratch.Element(Legacy + "Y")!.Attribute("F"));
        Assert.Equal("Label", edited.Element(Legacy + "Text")!.Value);

        Assert.True(shape?.RemoveShapeSheetSection("Scratch") ?? edge!.RemoveShapeSheetSection("Scratch"));
        Assert.Empty(XmlShape(package).Elements(Legacy + "Scratch"));
        var added = new VisioShapeSheetSection("Scratch");
        var row = added.GetOrAddRow("New"); row.Name = null; row.Index = 9;
        row.SetCell("X", "4");
        Set(added);
        scratch = Assert.Single(XmlShape(package).Elements(Legacy + "Scratch"));
        Assert.Equal("9", (string?)scratch.Attribute("IX"));
        Assert.Equal("4", scratch.Element(Legacy + "X")!.Value);

        void Set(VisioShapeSheetSection value) { if (shape != null) shape.SetShapeSheetSection(value); else edge!.SetShapeSheetSection(value); }
        XElement XmlShape(bool throughPackage) {
            var candidate = throughPackage ? VisioDocument.Load(new MemoryStream(document.ToBytes())) : document;
            return XDocument.Load(new MemoryStream(candidate.ToLegacyXmlResult().Value)).Descendants(Legacy + "Pages").Descendants(Legacy + "Shape").Single();
        }
    }

    [Theory]
    [InlineData("nxbre-ie3.vsx")]
    [InlineData("pronom-visio2002.vsx")]
    [InlineData("pronom-visio2002.vtx")]
    public void IndependentMasterFontOverridesPersistWithoutChangingTextAndFields(string file) {
        var document = VisioDocument.LoadLegacyXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", file)).Value;
        var before = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        var candidates = document.Masters.SelectMany(master => Tree(master.Shape).Select(shape => (master, shape)))
            .Where(pair => !string.IsNullOrWhiteSpace(pair.shape.Text)).ToArray();
        var selected = candidates.FirstOrDefault(pair => pair.shape.GetShapeSheetSections().Any(s => s.Name == "Character"));
        if (selected.shape == null) selected = candidates.First();
        var character = selected.shape.GetShapeSheetSections().SingleOrDefault(s => s.Name == "Character") ?? new VisioShapeSheetSection("Character");
        var row = character.Rows.FirstOrDefault() ?? character.GetOrAddRow("New");
        row.Name = null; row.Index ??= 0;
        row.SetCell("Size", "0.25", unit: "PT");
        selected.shape.SetShapeSheetSection(character);
        foreach (var candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            var xml = XDocument.Load(new MemoryStream(candidate.ToLegacyXmlResult().Value));
            var master = xml.Descendants(Legacy + "Master").Single(m => (string?)m.Attribute("NameU") == selected.master.NameU);
            var shape = master.Descendants(Legacy + "Shape").Single(s => (string?)s.Attribute("ID") == selected.shape.Id);
            var saved = Assert.Single(shape.Elements(Legacy + "Char"));
            Assert.Equal("0.25", saved.Element(Legacy + "Size")!.Value);
            Assert.Null(saved.Element(Legacy + "Size")!.Attribute("F"));
            Assert.Equal(TextAndFields(before), TextAndFields(xml));
        }
    }

    private static IEnumerable<VisioShape> Tree(VisioShape shape) {
        yield return shape;
        foreach (var child in shape.Children) foreach (var nested in Tree(child)) yield return nested;
    }

    private static string[] TextAndFields(XDocument xml) => xml.Descendants(Legacy + "Masters")
        .Descendants().Where(e => e.Name == Legacy + "Text" || e.Name == Legacy + "Field")
        .Select(e => e.ToString(SaveOptions.DisableFormatting)).ToArray();
}

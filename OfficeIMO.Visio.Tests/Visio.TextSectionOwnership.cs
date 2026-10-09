using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioTextSectionOwnershipTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void TypedEditsRetainExtraNativeCellsAndProducerOnlyStyleFlags(bool connector, bool package) {
        var document = Reopen(Load(connector), package);
        VisioTextStyle style = Style(document, connector);
        Assert.Equal(12, style.Size);
        Assert.True(style.Bold);
        style.Size = 17;
        style.Bold = false;
        style.Italic = true;
        style.HorizontalAlignment = VisioTextHorizontalAlignment.Right;

        document = Reopen(document, package);
        XElement shape = SavedShape(document);
        XElement character = Assert.Single(shape.Elements(Legacy + "Char"));
        Assert.Equal("7", (string?)character.Attribute("IX"));
        Assert.Equal("18", character.Element(Legacy + "Style")!.Value);
        Assert.Null(character.Element(Legacy + "Style")!.Attribute("F"));
        Assert.Null(character.Element(Legacy + "Style")!.Attribute("Err"));
        Assert.Equal("0.8", character.Element(Legacy + "FontScale")!.Value);
        Assert.Equal("GUARD(0.8)", (string?)character.Element(Legacy + "FontScale")!.Attribute("F"));
        Assert.Equal(17, Style(document, connector).Size);
        XElement paragraph = Assert.Single(shape.Elements(Legacy + "Para"));
        Assert.Equal("4", (string?)paragraph.Attribute("IX"));
        Assert.Equal("2", paragraph.Element(Legacy + "HorzAlign")!.Value);
        Assert.Equal("0.1", paragraph.Element(Legacy + "IndLeft")!.Value);

        style = Style(document, connector);
        style.Bold = null;
        style.Italic = null;
        style.SmallCaps = null;
        style.UnderlineStyle = null;
        document = Reopen(document, package);
        character = Assert.Single(SavedShape(document).Elements(Legacy + "Char"));
        Assert.Equal("16", character.Element(Legacy + "Style")!.Value);
        Assert.Equal("0.8", character.Element(Legacy + "FontScale")!.Value);
        Assert.Equal("Label", SavedShape(document).Element(Legacy + "Text")!.Value);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ComplexNativeTextRowsRemainUniqueWhenTextBlockFormattingIsEdited(bool connector, bool multipleRows) {
        string character = multipleRows
            ? "<Char IX='7'><Size>0.16666666666666667</Size></Char><Char IX='9'><Size>0.25</Size></Char>"
            : "<Char IX='7'><Case>9</Case><Size>0.16666666666666667</Size></Char>";
        string paragraph = multipleRows
            ? "<Para IX='4'><HorzAlign>0</HorzAlign></Para><Para IX='6'><HorzAlign>2</HorzAlign></Para>"
            : "<Para IX='4'><HorzAlign>99</HorzAlign></Para>";
        var document = Load(connector, character, paragraph);
        var original = SavedShape(document);
        VisioTextStyle style = Style(document, connector);
        style.LeftMargin = 0.2;
        style.Size = 24;
        style.HorizontalAlignment = VisioTextHorizontalAlignment.Center;
        foreach (var candidate in new[] { document, Reopen(document, false), Reopen(document, true) }) {
            XElement saved = SavedShape(candidate);
            Assert.Equal(original.Elements(Legacy + "Char").Select(e => e.ToString()),
                saved.Elements(Legacy + "Char").Select(e => e.ToString()));
            Assert.Equal(original.Elements(Legacy + "Para").Select(e => e.ToString()),
                saved.Elements(Legacy + "Para").Select(e => e.ToString()));
            Assert.Equal("0.2", saved.Element(Legacy + "TextBlock")!.Element(Legacy + "LeftMargin")!.Value);
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ShapeSheetTextRowsSynchronizeTypedEditsReplacementAndRemoval(bool connector, bool package) {
        var document = Reopen(Load(connector), package);
        VisioTextStyle style = Style(document, connector);
        style.LeftMargin = 0.2;
        VisioShapeSheetSection character = Sections(document, connector).Single(s => s.Name == "Character");
        character.Rows.Single().SetCell("Size", "0.25", unit: "PT");
        Set(document, connector, character);
        Assert.Equal(18, Style(document, connector).Size);
        Style(document, connector).Size = 24;
        Assert.Equal(24 / 72D, double.Parse(Sections(document, connector).Single(s => s.Name == "Character")
            .Rows.Single().FindCell("Size")!.Value!, System.Globalization.CultureInfo.InvariantCulture), 12);

        VisioShapeSheetSection paragraph = Sections(document, connector).Single(s => s.Name == "Paragraph");
        paragraph.Rows.Single().SetCell("HorzAlign", "2");
        Set(document, connector, paragraph);
        Assert.Equal(VisioTextHorizontalAlignment.Right, Style(document, connector).HorizontalAlignment);
        document = Reopen(document, package);
        Assert.Equal(24, Style(document, connector).Size);
        Assert.Single(SavedShape(document).Elements(Legacy + "Char"));
        Assert.Single(SavedShape(document).Elements(Legacy + "Para"));

        // Transfer a source-backed row to native-only ownership, then back.
        character = Sections(document, connector).Single(s => s.Name == "Character");
        VisioShapeSheetRow second = character.GetOrAddRow("second");
        second.Name = null;
        second.Index = 9;
        second.SetCell("Size", "0.5");
        Set(document, connector, character);
        Assert.Null(Style(document, connector).Size);
        Assert.Equal(2, Sections(document, connector).Single(s => s.Name == "Character").Rows.Count);
        VisioShapeSheetSection single = new("Char");
        VisioShapeSheetRow row = single.GetOrAddRow("single");
        row.Name = null;
        row.Index = 7;
        row.SetCell("Size", "0.25", unit: "PT");
        Set(document, connector, single);
        Assert.Equal(18, Style(document, connector).Size);
        Assert.Single(SavedShape(Reopen(document, package)).Elements(Legacy + "Char"));

        Assert.True(Remove(document, connector, "Character"));
        Assert.True(Remove(document, connector, "Para"));
        Assert.False(Remove(document, connector, "Character"));
        document = Reopen(document, package);
        Assert.Empty(SavedShape(document).Elements(Legacy + "Char"));
        Assert.Empty(SavedShape(document).Elements(Legacy + "Para"));
        Assert.Equal(0.2, Style(document, connector).LeftMargin);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ShapeSheetSizeEditKeepsPendingFontFamilyAssignment(bool connector) {
        var document = Load(connector);
        Style(document, connector).FontFamily = "Arial";
        var section = Sections(document, connector).Single(s => s.Name == "Character");
        section.Rows.Single().SetCell("Size", "0.25", unit: "PT");
        Set(document, connector, section);
        Assert.Equal("Arial", Style(document, connector).FontFamily);
        foreach (var candidate in new[] { Reopen(document, false), Reopen(document, true) }) {
            Assert.Equal("Arial", Style(candidate, connector).FontFamily);
            Assert.Equal(18, Style(candidate, connector).Size);
            Assert.Single(SavedShape(candidate).Elements(Legacy + "Char"));
        }

        Style(document, connector).FontFamily = "Courier New";
        section = Sections(document, connector).Single(s => s.Name == "Character");
        section.Rows.Single().SetCell("Font", "0");
        Set(document, connector, section);
        Assert.Equal("Calibri", Style(Reopen(document, true), connector).FontFamily);
    }

    [Fact]
    public void SparseInstanceParagraphEditRetainsMasterBulletCellsWithoutStyleBindings() {
        string source = $"<VisioDocument xmlns='{Legacy}'><FaceNames><FaceName ID='0' Name='Arial'/></FaceNames><Masters><Master ID='7' NameU='Bulleted'><Shapes><Shape ID='100'><XForm><Width>4</Width><Height>2</Height></XForm><Char IX='3'><Font>0</Font><Size>0.16666666666666667</Size></Char><Para IX='9'><HorzAlign>0</HorzAlign><Bullet>1</Bullet><BulletStr>*</BulletStr><IndLeft>0.25</IndLeft><IndFirst>-0.25</IndFirst></Para><Text><cp IX='3'/><pp IX='9'/>FIRST</Text></Shape></Shapes></Master></Masters><Pages><Page ID='0'><Shapes><Shape ID='1' Master='7'><XForm><PinX>3</PinX><PinY>2</PinY><Width>4</Width><Height>2</Height></XForm></Shape></Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
        XNamespace svg = "http://www.w3.org/2000/svg";
        Assert.Equal(new[] { "*", "FIRST" }, PaintedText(document));
        document.Pages[0].Shapes[0].TextStyle!.HorizontalAlignment = VisioTextHorizontalAlignment.Right;
        foreach (var candidate in new[] { document, Reopen(document, false), Reopen(document, true) }) {
            byte[] before = candidate.ToLegacyXmlResult().Value;
            Assert.Equal(new[] { "*", "FIRST" }, PaintedText(candidate));
            Assert.Equal(before, candidate.ToLegacyXmlResult().Value);
        }

        string[] PaintedText(VisioDocument candidate) => XDocument.Parse(candidate.Pages[0].ToSvg())
            .Descendants(svg + "text").Where(text => !string.IsNullOrWhiteSpace(text.Value)).Select(text => text.Value).ToArray();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RemovingAnAbsentTextSectionLeavesAuthoredFormattingAndSavedContentUnchanged(bool connector) {
        var document = VisioDocument.Create();
        var page = document.AddPage("Authored");
        if (connector) {
            page.AddConnector("1", new OfficeIMO.Drawing.OfficePoint(1, 1), new OfficeIMO.Drawing.OfficePoint(3, 1)).Label = "Label";
        } else {
            page.Shapes.Add(new VisioShape("1", 2, 1, 2, 1, "Label"));
        }
        VisioTextStyle style = Style(document, connector);
        style.FontFamily = "Arial";
        style.Size = 18;
        style.Bold = true;
        style.HorizontalAlignment = VisioTextHorizontalAlignment.Right;
        style.LeftMargin = 0.2;
        byte[] before = document.ToLegacyXmlResult().Value;
        Assert.Empty(Sections(document, connector));

        foreach (string name in new[] { "Character", "Char", "Paragraph", "Para" }) {
            Assert.False(Remove(document, connector, name));
            Assert.Same(style, Style(document, connector));
            Assert.Equal("Arial", style.FontFamily);
            Assert.Equal(18, style.Size);
            Assert.True(style.Bold);
            Assert.Equal(VisioTextHorizontalAlignment.Right, style.HorizontalAlignment);
            Assert.Equal(0.2, style.LeftMargin);
            Assert.Equal(before, document.ToLegacyXmlResult().Value);
        }
    }

    private static VisioDocument Load(bool connector, string? character = null, string? paragraph = null) {
        character ??= "<Char IX='7'><Font>0</Font><Style F='GUARD(17)' Err='#REF!'>17</Style><Size Unit='PT' F='GUARD(12 pt)'>0.16666666666666667</Size><FontScale F='GUARD(0.8)'>0.8</FontScale></Char>";
        paragraph ??= "<Para IX='4'><HorzAlign>0</HorzAlign><IndLeft>0.1</IndLeft></Para>";
        string oneD = connector ? "<XForm1D><BeginX>1</BeginX><BeginY>1</BeginY><EndX>3</EndX><EndY>1</EndY></XForm1D>" : "";
        string source = $"<VisioDocument xmlns='{Legacy}'><FaceNames><FaceName ID='0' Name='Calibri'/></FaceNames><Pages><Page ID='0'><Shapes><Shape ID='1' Type='Shape'><XForm><PinX>2</PinX><PinY>1</PinY><Width>2</Width><Height>1</Height></XForm>{oneD}{character}{paragraph}<Text><cp IX='7'/><pp IX='4'/>Label</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
    }

    private static VisioDocument Reopen(VisioDocument document, bool package) => package
        ? VisioDocument.Load(new MemoryStream(document.ToBytes()))
        : VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
    private static VisioTextStyle Style(VisioDocument document, bool connector) => connector
        ? document.Pages[0].Connectors[0].TextStyle ??= new VisioTextStyle()
        : document.Pages[0].Shapes[0].TextStyle ??= new VisioTextStyle();
    private static IReadOnlyList<VisioShapeSheetSection> Sections(VisioDocument document, bool connector) => connector
        ? document.Pages[0].Connectors[0].GetShapeSheetSections() : document.Pages[0].Shapes[0].GetShapeSheetSections();
    private static void Set(VisioDocument document, bool connector, VisioShapeSheetSection section) {
        if (connector) document.Pages[0].Connectors[0].SetShapeSheetSection(section);
        else document.Pages[0].Shapes[0].SetShapeSheetSection(section);
    }
    private static bool Remove(VisioDocument document, bool connector, string name) => connector
        ? document.Pages[0].Connectors[0].RemoveShapeSheetSection(name) : document.Pages[0].Shapes[0].RemoveShapeSheetSection(name);
    private static XElement SavedShape(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value))
        .Descendants(Legacy + "Pages").Descendants(Legacy + "Shape").Single();
}

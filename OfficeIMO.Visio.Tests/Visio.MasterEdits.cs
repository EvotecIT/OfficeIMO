using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioMasterEditsTests {
    private static readonly XNamespace Ns = "http://schemas.microsoft.com/visio/2003/core";

    [Fact]
    public void FontEditPreservesUntouchedCharacterCellFormula() {
        var document = LoadMaster("<Char IX='0'><Font>0</Font><Size F='GUARD(0.166666666666667)'>0.166666666666667</Size></Char><Text>label</Text>");
        document.Masters.First().Shape.TextStyle!.FontFamily = "Consolas";
        var xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        Assert.Equal("GUARD(0.166666666666667)", (string?)xml.Descendants(Ns + "Size").Single().Attribute("F"));
        Assert.Equal("Consolas", VisioDocument.Load(new MemoryStream(document.ToBytes())).Masters.First().Shape.TextStyle!.FontFamily);
    }

    [Fact]
    public void ConnectionEditPreservesOtherFormulasAndUnmodeledCells() {
        var document = LoadMaster("<Connection IX='0'><X F='Width*0.5'>1</X><Y F='Height*0.5'>1</Y><DirX>0</DirX><DirY>1</DirY><Type>2</Type><Prompt>Keep this</Prompt></Connection>");
        document.Masters.First().Shape.ConnectionPoints[0].X = 2;
        var xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        var row = xml.Descendants(Ns + "Connection").Single();
        Assert.Equal("2", row.Element(Ns + "X")!.Value);
        Assert.Null(row.Element(Ns + "X")!.Attribute("F"));
        Assert.Equal("Height*0.5", (string?)row.Element(Ns + "Y")!.Attribute("F"));
        Assert.Equal("2", row.Element(Ns + "Type")!.Value);
        Assert.Equal("Keep this", row.Element(Ns + "Prompt")!.Value);
    }

    [Fact]
    public void NewStringChildIdIsMappedWithoutChangingExistingShapeIds() {
        var document = LoadMaster("<Shapes><Shape ID='2'><Text>existing</Text></Shape></Shapes>");
        document.Masters.First().Shape.Children.Insert(0, new VisioShape("new-child", 1, 1, 1, 1, "inserted"));
        var xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        var shapes = xml.Descendants(Ns + "Shape").ToArray();
        Assert.All(shapes, shape => Assert.True(uint.TryParse((string?)shape.Attribute("ID"), out _)));
        Assert.Equal(shapes.Length, shapes.Select(shape => (string?)shape.Attribute("ID")).Distinct().Count());
        Assert.Equal("2", (string?)shapes.Single(shape => (string?)shape.Element(Ns + "Text") == "existing").Attribute("ID"));
        var reopened = VisioDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(new[] { "new-child", "2" }, reopened.Masters.First().Shape.Children.Select(shape => shape.Id));
        string persistedChildId = (string)shapes.Single(shape => (string?)shape.Element(Ns + "Text") == "inserted").Attribute("ID")!;
        reopened.Masters.First().Shape.Children.Insert(0, new VisioShape(persistedChildId, 1, 1, 1, 1, "another"));
        var again = XDocument.Load(new MemoryStream(reopened.ToLegacyXmlResult().Value));
        Assert.Equal((string?)shapes.Single(shape => (string?)shape.Element(Ns + "Text") == "inserted").Attribute("ID"),
            (string?)again.Descendants(Ns + "Shape").Single(shape => (string?)shape.Element(Ns + "Text") == "inserted").Attribute("ID"));
    }

    [Fact]
    public void InsertedNumericIdCannotCollideWithPreservedAdditionalRoot() {
        var document = LoadMaster("<Shapes><Shape ID='3'><Text>existing child</Text></Shape></Shapes>",
            "<Shape ID='2'><Text>additional root</Text></Shape>");
        document.Masters.First().Shape.Children.Insert(0, new VisioShape("2", 1, 1, 1, 1, "inserted"));
        var xml = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        var shapes = xml.Descendants(Ns + "Shape").ToArray();
        Assert.Equal(shapes.Length, shapes.Select(shape => (string?)shape.Attribute("ID")).Distinct().Count());
        Assert.Equal("2", (string?)shapes.Single(shape => (string?)shape.Element(Ns + "Text") == "additional root").Attribute("ID"));
    }

    [Fact]
    public void ExplicitZeroLocalPinCanBeEditedToTheDefaultCenter() {
        var document = LoadMaster("<XForm><Width>2</Width><Height>2</Height><LocPinX>0</LocPinX><LocPinY>0</LocPinY></XForm>");
        document.Masters.First().Shape.LocPinX = 1;
        document.Masters.First().Shape.LocPinY = 1;
        var reopened = VisioDocument.Load(new MemoryStream(document.ToBytes())).Masters.First().Shape;
        Assert.Equal(1, reopened.LocPinX); Assert.Equal(1, reopened.LocPinY);
    }

    [Fact]
    public void ExplicitZeroDimensionsCanBeEditedToTheDefaultSize() {
        var document = LoadMaster("<XForm><Width>0</Width><Height>0</Height></XForm>");
        document.Masters.First().Shape.Width = 1;
        document.Masters.First().Shape.Height = 1;
        var reopened = VisioDocument.Load(new MemoryStream(document.ToBytes())).Masters.First().Shape;
        Assert.Equal(1, reopened.Width); Assert.Equal(1, reopened.Height);
    }

    private static VisioDocument LoadMaster(string body, string additionalShapes = "") {
        string source = "<VisioDocument xmlns='" + Ns + "'><Fonts><FontEntry ID='0' Name='Arial' CharSet='0'/></Fonts><Masters><Master ID='1' NameU='Example'><Shapes><Shape ID='1'>" + body + "</Shape>" + additionalShapes + "</Shapes></Master></Masters></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source)), VisioPackageType.Stencil).Value;
    }
}

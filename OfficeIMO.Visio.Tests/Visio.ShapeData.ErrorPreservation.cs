using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioShapeDataErrorPreservationTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void FormulaEditsAndLiteralOverridesClearOnlyTheirPreviousError(bool connector, bool package) {
        XNamespace legacy = "http://schemas.microsoft.com/visio/2003/core";
        string oneD = connector ? "<XForm1D><BeginX>1</BeginX><BeginY>1</BeginY><EndX>3</EndX><EndY>1</EndY></XForm1D>" : "";
        string source = $"<VisioDocument xmlns='{legacy}'><Pages><Page ID='0'><Shapes><Shape ID='1'><XForm><Width>2</Width><Height>1</Height></XForm>{oneD}<Prop ID='0' NameU='Amount'><Value F='1/0' Err='#DIV/0!'>1</Value><Label F='Sheet.99!Width' Err='#REF!'>Amount</Label></Prop></Shape></Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
        var shape = connector ? null : document.Pages[0].Shapes[0];
        var edge = connector ? document.Pages[0].Connectors[0] : null;
        var row = shape?.FindShapeData("Amount") ?? edge!.FindShapeData("Amount")!;
        Assert.Equal("#DIV/0!", (string?)Cell("Value").Attribute("Err"));
        row.ValueFormula = "1";
        Assert.Equal("1", (string?)Cell("Value").Attribute("F"));
        Assert.Null(Cell("Value").Attribute("Err"));
        Assert.Equal("#REF!", (string?)Cell("Label").Attribute("Err"));
        row.LabelFormula = "\"Amount\"";
        Assert.Null(Cell("Label").Attribute("Err"));
        if (shape != null) shape.SetShapeData("Amount", "2"); else edge!.SetShapeData("Amount", "2");
        Assert.Equal("2", Cell("Value").Value);
        Assert.Null(Cell("Value").Attribute("F"));
        Assert.Null(Cell("Value").Attribute("Err"));

        XElement Cell(string name) {
            var candidate = package ? VisioDocument.Load(new MemoryStream(document.ToBytes())) : document;
            return XDocument.Load(new MemoryStream(candidate.ToLegacyXmlResult().Value)).Descendants(legacy + "Prop").Single().Element(legacy + name)!;
        }
    }
}

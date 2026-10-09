using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioNativeCellAssignmentTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitUserWritesReplaceNullConditionsWhileKeepingMatchingErrors(bool useMethod) {
        var document = Load();
        var shape = document.Pages[0].Shapes[0];
        var row = shape.FindUserCell("Nullable")!;
        if (useMethod) shape.SetUserCell(row.Name, row.Value, row.Unit, row.Formula, row.Prompt);
        else { row.Value = row.Value; row.Prompt = row.Prompt; }
        AssertRoundTrips(document, xml => {
            AssertCell(User(Shape(xml, "Source"), "Value"), null, "#VALUE!");
            AssertCell(User(Shape(xml, "Source"), "Prompt"), null, "producer prompt error");
            AssertCell(Data(Shape(xml, "Source"), "Value"), "null", "producer data error");
        });
    }

    [Theory]
    [InlineData("Value")]
    [InlineData("Label")]
    [InlineData("Prompt")]
    [InlineData("Format")]
    [InlineData("SortKey")]
    [InlineData("Calendar")]
    [InlineData("LangID")]
    public void ShapeDataAssignmentAffectsOnlyItsOwnStringCell(string cell) {
        var document = Load();
        var row = document.Pages[0].Shapes[0].FindShapeData("Nullable")!;
        switch (cell) {
            case "Value": row.Value = row.Value; break;
            case "Label": row.Label = row.Label; break;
            case "Prompt": row.Prompt = row.Prompt; break;
            case "Format": row.Format = row.Format; break;
            case "SortKey": row.SortKey = row.SortKey; break;
            case "Calendar": row.Calendar = row.Calendar; break;
            case "LangID": row.LangId = row.LangId; break;
        }
        AssertRoundTrips(document, xml => {
            foreach (string name in StringCells)
                AssertCell(Data(Shape(xml, "Source"), name), name == cell ? null : "null", "producer data error");
        });
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void InstalledDetachedAssignmentsAreObservedOnShapesAndConnectors(bool connector, bool setCell) {
        var document = Load();
        var shape = document.Pages[0].Shapes[0];
        var edge = document.Pages[0].Connectors[0];
        var section = (connector ? edge.GetShapeSheetSections() : shape.GetShapeSheetSections()).Single(s => s.Name == "Paragraph");
        var row = section.Rows[0];
        var cell = row.FindCell("BulletStr")!;
        if (setCell) row.SetCell(cell.Name, cell.Value, cell.Formula, cell.Unit);
        else cell.Value = cell.Value;
        // Detached edits have no effect until the section is installed.
        AssertCell(Shape(Export(document), connector ? "Edge" : "Source").Element(Legacy + "Para")!.Element(Legacy + "BulletStr")!, "null", "producer paragraph error");
        if (connector) edge.SetShapeSheetSection(section); else shape.SetShapeSheetSection(section);
        AssertRoundTrips(document, xml => {
            AssertCell(Shape(xml, connector ? "Edge" : "Source").Element(Legacy + "Para")!.Element(Legacy + "BulletStr")!, null, "producer paragraph error");
            AssertCell(Shape(xml, connector ? "Source" : "Edge").Element(Legacy + "Para")!.Element(Legacy + "BulletStr")!, "null", "producer paragraph error");
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AssignmentIntentSurvivesCopiesAndDoesNotLeakIntoUntouchedSiblings(bool beforeCopy) {
        var document = Load(); var page = document.Pages[0]; var source = page.Shapes[0];
        if (beforeCopy) Assign(source);
        var copy = page.DuplicateShapes(page.Shapes.ToArray(), new VisioShapeDuplicationOptions { OffsetX = 0, OffsetY = 0 })[0];
        copy.NameU = "Copy";
        if (!beforeCopy) Assign(copy);
        var second = page.DuplicateShapes(new[] { copy }, new VisioShapeDuplicationOptions { OffsetX = 0, OffsetY = 0 })[0];
        second.NameU = "Second";
        AssertRoundTrips(document, xml => {
            foreach (string name in new[] { "Copy", "Second" }) {
                AssertCell(User(Shape(xml, name), "Value"), null, "#VALUE!");
                AssertCell(Data(Shape(xml, name), "Value"), null, "producer data error");
                AssertCell(Data(Shape(xml, name), "Prompt"), "null", "producer data error");
            }
            AssertCell(User(Shape(xml, "Source"), "Value"), beforeCopy ? null : "null", "#VALUE!");
        });
        static void Assign(VisioShape shape) {
            var user = shape.FindUserCell("Nullable")!; user.Value = "temporary"; user.Value = "";
            var data = shape.FindShapeData("Nullable")!; data.Value = "temporary"; data.Value = "";
        }
    }

    [Fact]
    public void UneditedDetachedSectionsAndCopiedMasterInstancesKeepNativeNulls() {
        var document = Load(); var source = document.Pages[0].Shapes[0];
        source.SetShapeSheetSection(source.GetShapeSheetSections().Single(s => s.Name == "Paragraph"));
        var master = document.RegisterMaster("Null master", source, "0");
        var untouched = document.Pages[0].AddShape("untouched", master, 1, 1, 2, 1); untouched.NameU = "Untouched";
        var edited = document.Pages[0].AddShape("edited", master, 1, 1, 2, 1); edited.NameU = "Edited";
        edited.FindUserCell("Nullable")!.Prompt = "";
        AssertRoundTrips(document, xml => {
            foreach (string name in new[] { "Source", "Untouched" }) {
                AssertCell(User(Shape(xml, name), "Prompt"), "null", "producer prompt error");
                AssertCell(Shape(xml, name).Element(Legacy + "Para")!.Element(Legacy + "BulletStr")!, "null", "producer paragraph error");
            }
            AssertCell(User(Shape(xml, "Edited"), "Prompt"), null, "producer prompt error");
            AssertCell(xml.Descendants(Legacy + "Master").Single().Descendants(Legacy + "User").Single(r => (string?)r.Attribute("NameU") == "Nullable").Element(Legacy + "Prompt")!, "null", "producer prompt error");
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SetShapeDataWritesReplaceNullsOnBothTargetKinds(bool connector) {
        var document = Load();
        if (connector) document.Pages[0].Connectors[0].SetShapeData("Nullable", "", prompt: "");
        else document.Pages[0].Shapes[0].SetShapeData("Nullable", "", prompt: "");
        AssertRoundTrips(document, xml => {
            var shape = Shape(xml, connector ? "Edge" : "Source");
            var value = Data(shape, "Value");
            Assert.Null(value.Attribute("V")); Assert.Equal("", value.Value);
            Assert.Null(value.Attribute("F")); Assert.Null(value.Attribute("Err"));
            AssertCell(Data(shape, "Prompt"), null, "producer data error");
            AssertCell(Data(shape, "Label"), "null", "producer data error");
        });
    }

    [Fact]
    public void AuthoredReplacementRowsDoNotAcquireTheRemovedRowsNullCondition() {
        const string xml = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Pages><Page ID='0'><Shapes><Shape ID='1'>"
            + "<XForm><Width>2</Width><Height>1</Height></XForm><User ID='0' NameU='Nullable'><Value V='null'/></User>"
            + "<Prop ID='0' NameU='Nullable'><Value V='null'/></Prop></Shape></Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
        var shape = document.Pages[0].Shapes[0];
        shape.UserCells.Clear(); shape.UserCells.Add(new VisioUserCell("Nullable", ""));
        shape.ShapeData.Clear(); shape.ShapeData.Add(new VisioShapeDataRow("Nullable", ""));
        AssertRoundTrips(document, exported => {
            Assert.All(exported.Descendants(Legacy + "Value"), value => {
                Assert.Equal("", value.Value); Assert.Null(value.Attribute("V"));
            });
        });
    }

    private static readonly string[] StringCells = { "Value", "Label", "Prompt", "Format", "SortKey", "Calendar", "LangID" };
    private static VisioDocument Load() {
        string cells = "<Para IX='0'><BulletStr V='null' Unit='STR' F='Inh' Err='producer paragraph error'/></Para>"
            + "<User ID='2' NameU='Nullable'><Value V='null' Unit='STR' F='Inh' Err='#VALUE!'/><Prompt V='null' Unit='STR' F='Inh' Err='producer prompt error'/></User>"
            + "<Prop ID='3' NameU='Nullable'>" + string.Concat(StringCells.Select(name => $"<{name} V='null' Unit='STR' F='Inh' Err='producer data error'/>")) + "</Prop>";
        string xml = $"<VisioDocument xmlns='{Legacy}'><Pages><Page ID='0' Name='Source'><Shapes><Shape ID='1' NameU='Source'><XForm><Width>2</Width><Height>1</Height></XForm>{cells}</Shape>"
            + "<Shape ID='4' NameU='Other'><XForm><Width>1</Width><Height>1</Height></XForm></Shape>"
            + $"<Shape ID='5' NameU='Edge'><XForm1D><BeginX>0</BeginX><BeginY>0</BeginY><EndX>2</EndX><EndY>0</EndY></XForm1D>{cells}<Text>Edge</Text></Shape>"
            + "</Shapes><Connects><Connect FromSheet='5' FromCell='BeginX' ToSheet='1' ToCell='PinX'/><Connect FromSheet='5' FromCell='EndX' ToSheet='4' ToCell='PinX'/></Connects></Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
    }
    private static void AssertRoundTrips(VisioDocument document, Action<XDocument> assert) {
        assert(Export(document));
        var package = VisioDocument.Load(new MemoryStream(document.ToBytes())); assert(Export(package));
        var legacy = VisioDocument.LoadLegacyXml(new MemoryStream(package.ToLegacyXmlResult().Value)).Value; assert(Export(legacy));
    }
    private static XDocument Export(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
    private static XElement Shape(XDocument xml, string name) => xml.Descendants(Legacy + "Page").Descendants(Legacy + "Shape")
        .Single(s => name == "Edge" ? s.Element(Legacy + "Text")?.Value == name : (string?)s.Attribute("NameU") == name);
    private static XElement User(XElement shape, string cell) => shape.Elements(Legacy + "User").Single(r => (string?)r.Attribute("NameU") == "Nullable").Element(Legacy + cell)!;
    private static XElement Data(XElement shape, string cell) => shape.Elements(Legacy + "Prop").Single(r => (string?)r.Attribute("NameU") == "Nullable").Element(Legacy + cell)!;
    private static void AssertCell(XElement cell, string? condition, string error) {
        Assert.Equal("", cell.Value); Assert.Equal(condition, (string?)cell.Attribute("V"));
        Assert.Equal("Inh", (string?)cell.Attribute("F")); Assert.Equal("STR", (string?)cell.Attribute("Unit"));
        Assert.Equal(error, (string?)cell.Attribute("Err"));
    }
}

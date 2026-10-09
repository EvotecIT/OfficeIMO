using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioLegacyLayerTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Native = "http://schemas.microsoft.com/office/visio/2012/main";

    [Fact]
    public void UntouchedLayerRowsKeepSourceIdentityAndNativeCellState() {
        VisioDocument document = Load();
        foreach (VisioDocument candidate in new[] { document, Reopen(document) }) {
            XElement[] rows = Export(candidate).Descendants(Legacy + "PageSheet").Single().Elements(Legacy + "Layer").ToArray();
            Assert.Equal(new[] { "0", "5" }, rows.Select(row => (string?)row.Attribute("IX")));
            Assert.Equal("null", (string?)rows[0].Element(Legacy + "NameUniv")!.Attribute("V"));
            Assert.Equal("producer error", (string?)rows[0].Element(Legacy + "ColorTrans")!.Attribute("Err"));
            Assert.Equal("GUARD(FALSE)", (string?)rows[1].Element(Legacy + "Visible")!.Attribute("F"));
            Assert.Contains("Hidden", candidate.Pages[0].Shapes[0].LayerNames);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitLayerEditsReplaceFormulasAndProducerStateEvenForSameValues(bool sameValue) {
        VisioDocument document = Load();
        VisioLayer first = document.Pages[0].Layers[0], hidden = document.Pages[0].Layers[1];
        first.NameU = first.NameU;
        first.ColorTransparency = sameValue ? first.ColorTransparency : 50;
        hidden.Visible = sameValue ? hidden.Visible : true;
        foreach (VisioDocument candidate in new[] { document, Reopen(document) }) {
            XElement[] rows = Export(candidate).Descendants(Legacy + "PageSheet").Single().Elements(Legacy + "Layer").ToArray();
            Assert.Null(rows[0].Element(Legacy + "NameUniv")!.Attribute("V"));
            XElement transparency = rows[0].Element(Legacy + "ColorTrans")!;
            Assert.Equal(sameValue ? "0" : "50", transparency.Value);
            Assert.Null(transparency.Attribute("F"));
            Assert.Null(transparency.Attribute("Err"));
            XElement visible = rows[1].Element(Legacy + "Visible")!;
            Assert.Equal(sameValue ? "0" : "1", visible.Value);
            Assert.Null(visible.Attribute("F"));
            Assert.Null(visible.Attribute("Err"));
            Assert.Equal("BOOL", (string?)visible.Attribute("Unit"));
            Assert.Equal("GUARD(FALSE)", (string?)rows[1].Element(Legacy + "Print")!.Attribute("F"));
        }
    }

    [Fact]
    public void DuplicatedAndReorderedLayersKeepRowIdentityAndOrdinalMembership() {
        VisioDocument document = Load();
        VisioPage copy = document.DuplicatePage(document.Pages[0], "Copy");
        VisioLayer hidden = copy.Layers[1];
        copy.Layers.RemoveAt(1);
        copy.Layers.Insert(0, hidden);
        copy.AddLayer("New");
        copy.AddLayer("Another");
        copy.AddLayer("Fifth row");
        copy.AddLayer("Row collision");
        foreach (VisioDocument candidate in new[] { document, Reopen(document) }) {
            XElement page = Export(candidate).Descendants(Legacy + "Page").Single(row => (string?)row.Attribute("Name") == "Copy");
            XElement[] rows = page.Element(Legacy + "PageSheet")!.Elements(Legacy + "Layer").ToArray();
            Assert.Equal(new[] { "5", "0", "2", "3", "4", "6" }, rows.Select(row => (string?)row.Attribute("IX")));
            Assert.Equal("0", page.Descendants(Legacy + "LayerMember").Single().Value);
            Assert.Contains("Hidden", candidate.Pages.Single(p => p.Name == "Copy").Shapes[0].LayerNames);
            Assert.Equal("producer error", (string?)rows[1].Element(Legacy + "ColorTrans")!.Attribute("Err"));
            Assert.Equal("null", (string?)rows[1].Element(Legacy + "NameUniv")!.Attribute("V"));
        }
    }

    [Theory]
    [InlineData("unchanged")]
    [InlineData("reorder")]
    [InlineData("clear")]
    public void LayerMembershipKeepsSourceStateUntilMembershipOrOrdinalPositionsChange(string edit) {
        const string membership = "<LayerMem><LayerMember Unit='STR' F='GUARD(&quot;1&quot;)' Err='producer membership error'>1</LayerMember></LayerMem>";
        string source = "<VisioDocument xmlns='" + Legacy + "'><Pages><Page ID='0' Name='Members'><PageSheet>"
            + "<Layer IX='0'><Name>First</Name></Layer><Layer IX='5'><Name>Second</Name></Layer></PageSheet><Shapes>"
            + "<Shape ID='1' Type='Group'><XForm><Width>2</Width><Height>2</Height></XForm><Shapes>"
            + "<Shape ID='2'><XForm><Width>1</Width><Height>1</Height></XForm>" + membership + "<Text>Nested</Text></Shape>"
            + "</Shapes></Shape><Shape ID='3'><XForm1D><BeginX>0</BeginX><BeginY>0</BeginY><EndX>2</EndX><EndY>2</EndY></XForm1D>"
            + membership + "<Text>Free route</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        VisioDocument document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
        VisioPage page = document.Pages[0];
        if (edit == "clear") {
            page.Shapes[0].Children[0].LayerNames.Clear();
            page.Connectors[0].LayerNames.Clear();
        } else if (edit == "reorder") {
            VisioLayer second = page.Layers[1]; page.Layers.RemoveAt(1); page.Layers.Insert(0, second);
        }
        document.DuplicatePage(page, "Copy");
        foreach (VisioDocument candidate in new[] { document, Reopen(document) }) {
            XElement[] members = Export(candidate).Descendants(Legacy + "LayerMember").ToArray();
            Assert.Equal(4, members.Length);
            foreach (XElement member in members) {
                Assert.Equal(edit == "clear" ? "" : edit == "reorder" ? "0" : "1", member.Value);
                Assert.Equal("STR", (string?)member.Attribute("Unit"));
                Assert.Equal(edit == "unchanged" ? "GUARD(\"1\")" : null, (string?)member.Attribute("F"));
                Assert.Equal(edit == "unchanged" ? "producer membership error" : null, (string?)member.Attribute("Err"));
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RestoringMembershipAfterSavingAnEditDoesNotResurrectProducerState(bool hasFormula) {
        string member = "<LayerMem><LayerMember Unit='STR'" + (hasFormula ? " F='GUARD(&quot;0&quot;)'" : "")
            + " Err='producer membership error'>0</LayerMember></LayerMem>";
        string source = "<VisioDocument xmlns='" + Legacy + "'><Pages><Page ID='0'><PageSheet><Layer IX='0'><Name>Review</Name></Layer></PageSheet><Shapes>"
            + "<Shape ID='1' Type='Group'><XForm><Width>2</Width><Height>2</Height></XForm>" + member + "<Shapes>"
            + "<Shape ID='2'><XForm><Width>1</Width><Height>1</Height></XForm>" + member + "</Shape></Shapes></Shape>"
            + "<Shape ID='3'><XForm1D><BeginX>0</BeginX><BeginY>0</BeginY><EndX>2</EndX><EndY>2</EndY></XForm1D>" + member + "</Shape>"
            + "</Shapes></Page></Pages></VisioDocument>";
        VisioDocument document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
        VisioPage page = document.Pages[0];
        foreach (VisioShape shape in page.AllShapes()) shape.LayerNames.Clear();
        page.Connectors[0].LayerNames.Clear();
        Assert.All(Export(document).Descendants(Legacy + "LayerMember"), cell => Assert.Empty(cell.Value));
        foreach (VisioShape shape in page.AllShapes()) shape.LayerNames.Add("Review");
        page.Connectors[0].LayerNames.Add("Review");
        document.DuplicatePage(page, "Restored copy");
        foreach (VisioDocument candidate in new[] { document, Reopen(document) }) {
            XElement[] members = Export(candidate).Descendants(Legacy + "LayerMember").ToArray();
            Assert.Equal(6, members.Length);
            foreach (XElement cell in members) {
                Assert.Equal("0", cell.Value);
                Assert.Equal("STR", (string?)cell.Attribute("Unit"));
                Assert.Null(cell.Attribute("F"));
                Assert.Null(cell.Attribute("Err"));
            }
        }
    }

    [Fact]
    public void RepeatedPageCopiesRemapLayerFormulasAndKeepNativeProducerState() {
        string source = "<VisioDocument xmlns='" + Legacy + "'><Pages><Page ID='0' Name='Source'><PageSheet><Layer IX='0'>"
            + "<Name>Review</Name><Visible F='Sheet.1!User.Shown' Err='producer layer error'>1</Visible>"
            + "<ColorTrans F='Sheet.1!User.Transparency' Err='#VALUE!'>0</ColorTrans></Layer></PageSheet><Shapes><Shape ID='1'>"
            + "<XForm><Width>2</Width><Height>1</Height></XForm><User ID='1' NameU='Shown'><Value>1</Value></User>"
            + "<User ID='2' NameU='Transparency'><Value>0</Value></User><Text>Source</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        VisioDocument document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
        VisioPage first = document.DuplicatePage(document.Pages[0], "Copy");
        document.DuplicatePage(first, "Second");
        foreach (VisioDocument candidate in new[] { document, Reopen(document) }) {
            foreach (XElement page in Export(candidate).Descendants(Legacy + "Page")) {
                string? shapeId = (string?)page.Descendants(Legacy + "Shape").Single().Attribute("ID");
                XElement layer = page.Element(Legacy + "PageSheet")!.Element(Legacy + "Layer")!;
                Assert.Equal("Sheet." + shapeId + "!User.Shown", (string?)layer.Element(Legacy + "Visible")!.Attribute("F"));
                Assert.Equal("producer layer error", (string?)layer.Element(Legacy + "Visible")!.Attribute("Err"));
                Assert.Equal("Sheet." + shapeId + "!User.Transparency", (string?)layer.Element(Legacy + "ColorTrans")!.Attribute("F"));
                Assert.Equal("#VALUE!", (string?)layer.Element(Legacy + "ColorTrans")!.Attribute("Err"));
            }
        }
    }

    [Theory]
    [InlineData("angularjs-concepts.vdx")]
    [InlineData("nxbre-chocolatebox.vdx")]
    public void IndependentProducerPageLayersKeepIdentityPropertiesAndMembership(string file) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", file);
        XDocument original = XDocument.Load(path);
        VisioDocument document = VisioDocument.LoadLegacyXml(path).Value;
        foreach (VisioDocument candidate in new[] { document, Reopen(document) }) {
            XDocument saved = Export(candidate);
            foreach (XElement sourcePage in original.Descendants(Legacy + "Pages").Elements(Legacy + "Page")) {
                XElement page = saved.Descendants(Legacy + "Pages").Elements(Legacy + "Page")
                    .Single(p => (string?)p.Attribute("NameU") == (string?)sourcePage.Attribute("NameU"));
                XElement[] expected = sourcePage.Element(Legacy + "PageSheet")!.Elements(Legacy + "Layer").ToArray();
                XElement[] actual = page.Element(Legacy + "PageSheet")!.Elements(Legacy + "Layer").ToArray();
                Assert.Equal(expected.Length, actual.Length);
                for (int i = 0; i < expected.Length; i++) {
                    Assert.Equal((string?)expected[i].Attribute("IX"), (string?)actual[i].Attribute("IX"));
                    foreach (XElement cell in expected[i].Elements()) {
                        XElement emitted = actual[i].Element(cell.Name)!;
                        Assert.Equal(cell.Value, emitted.Value);
                        foreach (XAttribute attribute in cell.Attributes()) Assert.Equal(attribute.Value, (string?)emitted.Attribute(attribute.Name));
                    }
                }
                foreach (XElement sourceShape in sourcePage.Descendants(Legacy + "Shape").Where(s => s.Element(Legacy + "LayerMem") != null)) {
                    XElement shape = page.Descendants(Legacy + "Shape").Single(s => (string?)s.Attribute("ID") == (string?)sourceShape.Attribute("ID"));
                    Assert.Equal(sourceShape.Element(Legacy + "LayerMem")!.Element(Legacy + "LayerMember")!.Value,
                        shape.Element(Legacy + "LayerMem")!.Element(Legacy + "LayerMember")!.Value);
                }
            }
        }
    }

    private static VisioDocument Load() {
        string source = "<VisioDocument xmlns='" + Legacy + "'><Pages><Page ID='0' Name='Layers'><PageSheet>"
            + "<Layer IX='0'><Name>Visible</Name><NameUniv V='null'>Visible</NameUniv><ColorTrans F='Custom' Err='producer error'>0</ColorTrans></Layer>"
            + "<Layer IX='5'><Name>Hidden</Name><Visible Unit='BOOL' F='GUARD(FALSE)' Err='#VALUE!'>0</Visible><Print F='GUARD(FALSE)'>0</Print></Layer>"
            + "</PageSheet><Shapes><Shape ID='1'><XForm><PinX>2</PinX><PinY>2</PinY><Width>2</Width><Height>1</Height></XForm>"
            + "<LayerMem><LayerMember>1</LayerMember></LayerMem><Text>Layer member</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
    }

    private static VisioDocument Reopen(VisioDocument document) => VisioDocument.Load(new MemoryStream(document.ToBytes()));
    private static XDocument Export(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
}

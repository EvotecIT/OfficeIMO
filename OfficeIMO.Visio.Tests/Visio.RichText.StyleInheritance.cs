using System.Globalization;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioRichTextStyleInheritanceTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void IndependentBulletLabelInheritsArialEightPointTextWithoutRewritingSource(int reopen) {
        VisioDocument document = Reopen(VisioDocument.LoadLegacyXml(Path.Combine(AppContext.BaseDirectory,
            "Fixtures", "LegacyXml", "shorewall-netfilter.vdx")).Value, reopen);
        byte[] before = document.ToLegacyXmlResult().Value;
        VisioPage page = document.Pages[0];
        VisioShape shape = page.Shapes.Single(value => value.Id == "32");
        VisioRichTextProjection projection = Assert.IsType<VisioRichTextProjection>(
            VisioRichTextProjection.Create(page, shape, 72, CancellationToken.None));
        Assert.Equal(new[] { "Raw", "Mangle", "Nat" }, projection.Paragraphs.Select(paragraph =>
            string.Concat(paragraph.Runs.Select(run => run.Text))));
        Assert.All(projection.Paragraphs, paragraph => Assert.Equal("•", paragraph.Label!.Run.Text));
        Assert.All(projection.Runs, run => {
            Assert.Equal("Arial", run.FontFamily);
            Assert.Equal(8D, run.FontSize, 8);
        });
        _ = page.ToSvg();
        Assert.Equal("Raw\nMangle\nNat\n", shape.Text);
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SparseLocalRowsUseTheirStyleChainAndAuthoritativeInheritedCaches(bool connector) {
        VisioDocument document = Load(ChainedStyles, Element(connector, "TextStyle='6'",
            "<Char IX='0'><Size Unit='PT' F='Inh'>0.2777777777777778</Size></Char>"
            + "<Char IX='1'><Color>#2244cc</Color></Char>", "<cp IX='0'/>AA<cp IX='1'/>BB"));
        foreach (VisioDocument candidate in Candidates(document)) {
            VisioRichTextProjection projection = Project(candidate, connector);
            OfficeRichTextRun first = Assert.Single(projection.Runs, run => run.Text == "AA");
            OfficeRichTextRun second = Assert.Single(projection.Runs, run => run.Text == "BB");
            AssertRun(first, 20, OfficeColor.FromRgb(204, 17, 34));
            AssertRun(second, 12, OfficeColor.FromRgb(34, 68, 204));
            Assert.True(first.Bold);
            Assert.True(second.Bold);
        }
    }

    [Fact]
    public void ExplicitInstanceStyleSelectsPropertiesInsteadOfMasterFormatting() {
        const string master = "<Masters><Master ID='8' NameU='Native text'><Shapes><Shape ID='1'>"
            + ShapeTransform + "<Char IX='0'><Font>4</Font><Color>#993300</Color>"
            + "<Size Unit='PT'>0.4166666666666667</Size></Char><Text>Master</Text></Shape></Shapes></Master></Masters>";
        VisioDocument document = Load(ChainedStyles + master,
            Element(false, "Master='8' TextStyle='6'", "", "<cp IX='0'/>Instance"));
        foreach (VisioDocument candidate in Candidates(document)) {
            OfficeRichTextRun run = Assert.Single(Project(candidate, false).Runs);
            Assert.Equal("Instance", run.Text);
            AssertRun(run, 12, OfficeColor.FromRgb(204, 17, 34));
            Assert.True(run.Bold);
        }
    }

    [Fact]
    public void ExplicitBaseStyleDoesNotReuseMasterCharacterOrParagraphCaches() {
        const string master = "<Masters><Master ID='8' NameU='Formatted master'><Shapes><Shape ID='1'>"
            + ShapeTransform + "<Char IX='0'><Font>4</Font><Color>#993300</Color><Size Unit='PT'>0.25</Size></Char>"
            + "<Char IX='9'><Font>4</Font><Color>#993300</Color><Size Unit='PT'>0.4166666666666667</Size></Char>"
            + "<Para IX='0'><HorzAlign>1</HorzAlign></Para><Para IX='9'><HorzAlign>2</HorzAlign></Para>"
            + "<Text><cp IX='9'/><pp IX='9'/>Master</Text></Shape></Shapes></Master></Masters>";
        VisioDocument document = Load("<StyleSheets><StyleSheet ID='0'/></StyleSheets>" + master,
            Element(false, "Master='8' TextStyle='0'", "", "<cp IX='9'/><pp IX='9'/>Instance"));
        foreach (VisioDocument candidate in Candidates(document)) {
            byte[] before = candidate.ToLegacyXmlResult().Value;
            VisioShape shape = Assert.Single(candidate.Pages[0].Shapes);
            Assert.Null(shape.TextStyle?.Size);
            Assert.Null(shape.TextStyle?.HorizontalAlignment);
            Assert.Equal("0", shape.NativeStyleReferences?.TextStyle);
            XElement text = Assert.Single(LabelSvg(candidate, false).Descendants(Svg + "text"));
            Assert.Equal("Instance", text.Value);
            Assert.InRange(double.Parse((string)text.Attribute("font-size")!, CultureInfo.InvariantCulture), 13.332, 13.334);
            XDocument native = XDocument.Load(new MemoryStream(candidate.ToLegacyXmlResult().Value));
            Assert.Equal("0", (string?)Assert.Single(native.Descendants(Legacy + "Page")
                .Descendants(Legacy + "Shape")).Attribute("TextStyle"));
            _ = candidate.Pages[0].ToPng(new VisioPngSaveOptions { Supersampling = 1 });
            Assert.Equal(before, candidate.ToLegacyXmlResult().Value);
        }
    }

    [Fact]
    public void MasterRowsMatchNativeIndicesAndUseFirstNonzeroRowForAnUnmatchedMarker() {
        const string master = "<Masters><Master ID='8' NameU='Indexed text'><Shapes><Shape ID='1'>"
            + ShapeTransform + "<Char IX='5'><Font>4</Font><Color>#cc1122</Color><Size Unit='PT'>0.25</Size></Char>"
            + "<Char IX='9'><Font>4</Font><Color>#008844</Color><Size Unit='PT'>0.3333333333333333</Size></Char>"
            + "<Text><cp IX='5'/>Master</Text></Shape></Shapes></Master></Masters>";
        VisioDocument document = Load(master,
            Element(false, "Master='8'", "", "<cp IX='9'/>Exact<cp IX='11'/>Fallback"));
        foreach (VisioDocument candidate in Candidates(document)) {
            VisioRichTextProjection projection = Project(candidate, false);
            AssertRun(Assert.Single(projection.Runs, run => run.Text == "Exact"), 24, OfficeColor.FromRgb(0, 136, 68));
            AssertRun(Assert.Single(projection.Runs, run => run.Text == "Fallback"), 18, OfficeColor.FromRgb(204, 17, 34));
        }
    }

    [Fact]
    public void PlainInstanceTextUsesTheFirstNativeMasterRowsInsteadOfSyntheticZeroRows() {
        const string master = "<Masters><Master ID='8' NameU='Nonzero first'><Shapes><Shape ID='1'>"
            + ShapeTransform + "<Char IX='5'><Font>4</Font><Color>#cc1122</Color><Size Unit='PT'>0.25</Size></Char>"
            + "<Para IX='5'><HorzAlign>2</HorzAlign></Para><Text><cp IX='5'/><pp IX='5'/>Master</Text>"
            + "</Shape></Shapes></Master></Masters>";
        VisioDocument document = Load(master, Element(false, "Master='8'", "", "Instance"));
        foreach (VisioDocument candidate in Candidates(document)) {
            byte[] before = candidate.ToLegacyXmlResult().Value;
            VisioRichTextProjection projection = Project(candidate, false);
            AssertRun(Assert.Single(projection.Runs), 18, OfficeColor.FromRgb(204, 17, 34));
            Assert.Equal(OfficeTextAlignment.Right, Assert.Single(projection.Paragraphs).Alignment);
            XElement text = Assert.Single(LabelSvg(candidate, false).Descendants(Svg + "text"));
            Assert.Equal("24", (string?)text.Attribute("font-size"));
            _ = candidate.Pages[0].ToPng(new VisioPngSaveOptions { Supersampling = 1 });
            Assert.Equal(before, candidate.ToLegacyXmlResult().Value);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ModeledAndSourceSectionEditsReachInheritedRunsRenderingAndReopen(bool connector) {
        VisioDocument modeled = Load(ChainedStyles, Element(connector, "TextStyle='6'",
            "<Char IX='4'><Color>#112233</Color><Size Unit='PT'>0.25</Size></Char>", "<cp IX='4'/>AA"));
        VisioPage page = modeled.Pages[0];
        VisioTextStyle style = connector ? page.Connectors[0].TextStyle! : page.Shapes[0].TextStyle!;
        style.Size = 22;
        foreach (VisioDocument candidate in Candidates(modeled)) {
            AssertRun(Assert.Single(Project(candidate, connector).Runs), 22, OfficeColor.FromRgb(17, 34, 51));
            XElement label = LabelSvg(candidate, connector);
            XElement text = Assert.Single(label.Descendants(Svg + "text"));
            Assert.Equal("Arial", (string?)text.Attribute("font-family"));
            Assert.InRange(double.Parse((string)text.Attribute("font-size")!, CultureInfo.InvariantCulture), 29.332, 29.334);
        }

        VisioDocument source = Load(ChainedStyles, Element(connector, "TextStyle='6'",
            "<Char IX='0'><Size Unit='PT'>0.2777777777777778</Size></Char>"
            + "<Char IX='1'><Color>#008844</Color></Char>", "<cp IX='0'/>AA<cp IX='1'/>BB"));
        page = source.Pages[0];
        VisioShapeSheetSection section = (connector ? page.Connectors[0].GetShapeSheetSections() : page.Shapes[0].GetShapeSheetSections())
            .Single(value => value.Name is "Char" or "Character");
        section.Rows.Single(row => row.Index == 1).FindCell("Color")!.Value = "#2244cc";
        if (connector) page.Connectors[0].SetShapeSheetSection(section); else page.Shapes[0].SetShapeSheetSection(section);
        foreach (VisioDocument candidate in Candidates(source)) {
            AssertRun(Assert.Single(Project(candidate, connector).Runs, run => run.Text == "BB"), 12, OfficeColor.FromRgb(34, 68, 204));
            XElement label = LabelSvg(candidate, connector);
            Assert.Equal("BB", string.Concat(label.Descendants(Svg + "text").Where(text =>
                string.Equals((string?)text.Attribute("fill"), "#2244cc", StringComparison.OrdinalIgnoreCase)).Select(text => text.Value)));
        }
    }

    private const string ShapeTransform = "<XForm><PinX>3.5</PinX><PinY>2.5</PinY><Width>5</Width><Height>2</Height></XForm>";
    private const string ChainedStyles = "<StyleSheets><StyleSheet ID='7'><Char IX='0'><Font>4</Font><Color>#123456</Color><Style>1</Style>"
        + "<Size Unit='PT'>0.1666666666666667</Size></Char></StyleSheet><StyleSheet ID='6' TextStyle='7'><Char IX='0'>"
        + "<Color>#cc1122</Color></Char></StyleSheet></StyleSheets>";
    private static string Element(bool connector, string attributes, string properties, string text) =>
        "<Shape ID='1' " + attributes + ">" + (connector
            ? "<XForm1D><BeginX>1</BeginX><BeginY>2.5</BeginY><EndX>6</EndX><EndY>2.5</EndY></XForm1D>"
                + "<TextXForm><TxtWidth>5</TxtWidth><TxtHeight>2</TxtHeight></TextXForm>"
            : ShapeTransform) + properties + "<Text>" + text + "</Text></Shape>";
    private static VisioDocument Load(string documentContent, string element) => VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(
        "<VisioDocument xmlns='" + Legacy + "'><FaceNames><FaceName ID='4' Name='Arial'/></FaceNames>" + documentContent
        + "<Pages><Page ID='0' Name='Text'><PageSheet><PageProps><PageWidth>7</PageWidth><PageHeight>5</PageHeight></PageProps></PageSheet>"
        + "<Shapes>" + element + "</Shapes></Page></Pages></VisioDocument>"))).Value;
    private static VisioDocument Reopen(VisioDocument document, int mode) => mode switch {
        1 => VisioDocument.Load(new MemoryStream(document.ToBytes())),
        2 => VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value,
        _ => document
    };
    private static IEnumerable<VisioDocument> Candidates(VisioDocument document) {
        yield return document;
        yield return Reopen(document, 1);
        yield return Reopen(document, 2);
    }
    private static VisioRichTextProjection Project(VisioDocument document, bool connector) {
        VisioPage page = document.Pages[0];
        return Assert.IsType<VisioRichTextProjection>(connector
            ? VisioRichTextProjection.Create(page, Assert.Single(page.Connectors), 72, CancellationToken.None)
            : VisioRichTextProjection.Create(page, Assert.Single(page.Shapes), 72, CancellationToken.None));
    }
    private static XElement LabelSvg(VisioDocument document, bool connector) => XDocument.Parse(document.Pages[0].ToSvg())
        .Descendants(Svg + "g").Single(group => group.Attribute(connector ? "data-visio-connector-id" : "data-visio-shape-id") != null);
    private static void AssertRun(OfficeRichTextRun run, double size, OfficeColor color) {
        Assert.Equal("Arial", run.FontFamily);
        Assert.Equal(size, run.FontSize, 8);
        Assert.Equal(color, run.Color);
    }
}

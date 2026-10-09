using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioTextBackgroundTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";

    [Fact]
    public void IndependentNxbreCalloutHasNoBackgroundAfterXmlAndPackageReopening() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "nxbre-chocolatebox.vdx");
        var imported = VisioDocument.LoadLegacyXml(path);
        Assert.DoesNotContain(imported.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "VDX_TEXT_BACKGROUND");
        foreach (VisioDocument candidate in ReopenedCandidates(imported.Value)) {
            var page = candidate.Pages.Single(page => page.Name == "Rules");
            VisioConnector callout = page.Connectors.Single(connector => connector.Id == "18");
            Assert.Equal(0, callout.TextStyle!.BackgroundColor!.Value.A);
            Assert.Equal(0, callout.TextStyle.BackgroundTransparency);
            XElement exported = Export(candidate).Descendants(Legacy + "Page").Single(page => (string?)page.Attribute("Name") == "Rules")
                .Descendants(Legacy + "Shape").Single(shape => shape.Element(Legacy + "Text")?.Value == callout.Label);
            XElement block = exported.Element(Legacy + "TextBlock")!;
            Assert.Equal("0", block.Element(Legacy + "TextBkgnd")!.Value);
            Assert.Equal("0", block.Element(Legacy + "TextBkgndTrans")!.Value);
        }
    }

    [Theory]
    [InlineData("0", 0, 0, 0, 0)]
    [InlineData("255", 0, 0, 0, 0)]
    [InlineData("1", 255, 0, 0, 0)]
    [InlineData("2", 255, 255, 255, 255)]
    [InlineData("3", 255, 18, 52, 86)]
    [InlineData("#AABBCC", 255, 170, 187, 204)]
    [InlineData("RGB(170,187,204)+1", 255, 170, 187, 204)]
    public void NativeBackgroundCachesResolveAndKeepTheirCellsForShapesAndConnectors(string value, int alpha, int red, int green, int blue) {
        VisioDocument document = Load(value);
        foreach (VisioDocument candidate in ReopenedCandidates(document)) {
            foreach (VisioTextStyle style in Styles(candidate)) {
                var color = style.BackgroundColor!.Value;
                Assert.Equal(alpha, color.A);
                if (alpha != 0) {
                    Assert.Equal(red, color.R); Assert.Equal(green, color.G); Assert.Equal(blue, color.B);
                }
                Assert.Equal(25, style.BackgroundTransparency);
            }
            foreach (XElement block in Export(candidate).Descendants(Legacy + "TextBlock")) {
                AssertCell(block.Element(Legacy + "TextBkgnd")!, value, "GUARD(" + value + ")", "DL");
                AssertCell(block.Element(Legacy + "TextBkgndTrans")!, "0.2500", "GUARD(25%)", "%");
            }
        }
    }

    [Fact]
    public void ExplicitSameValueAssignmentsReplaceNativeFormulasAndErrors() {
        VisioDocument document = Load("0", error: "#NUM!");
        foreach (VisioTextStyle style in Styles(document)) {
            style.BackgroundColor = style.BackgroundColor;
            style.BackgroundTransparency = style.BackgroundTransparency;
        }
        foreach (VisioDocument candidate in ReopenedCandidates(document)) {
            foreach (XElement block in Export(candidate).Descendants(Legacy + "TextBlock")) {
                AssertCell(block.Element(Legacy + "TextBkgnd")!, "0", null, null);
                AssertCell(block.Element(Legacy + "TextBkgndTrans")!, "0.25", null, null);
                Assert.Null(block.Element(Legacy + "TextBkgnd")!.Attribute("Err"));
                Assert.Null(block.Element(Legacy + "TextBkgndTrans")!.Attribute("Err"));
            }
        }
    }

    [Theory]
    [InlineData("#NUM!")]
    [InlineData("producer-specific error")]
    public void ExplicitSameValueAssignmentsClearErrorsWhenCanonicalCacheIsUnchanged(string error) {
        XDocument source = Export(Load("0", error: error));
        foreach (XElement cell in source.Descendants(Legacy + "TextBkgnd").Concat(source.Descendants(Legacy + "TextBkgndTrans"))) {
            cell.Attribute("F")?.Remove(); cell.Attribute("Unit")?.Remove();
            if (cell.Name.LocalName == "TextBkgndTrans") cell.Value = "0.25";
        }
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source.ToString()))).Value;
        Assert.All(Export(document).Descendants(Legacy + "TextBkgnd"), cell => Assert.Equal(error, (string?)cell.Attribute("Err")));
        foreach (VisioTextStyle style in Styles(document)) {
            style.BackgroundColor = style.BackgroundColor;
            style.BackgroundTransparency = style.BackgroundTransparency;
        }
        foreach (VisioDocument candidate in ReopenedCandidates(document))
            Assert.All(Export(candidate).Descendants(Legacy + "TextBkgnd").Concat(Export(candidate).Descendants(Legacy + "TextBkgndTrans")),
                cell => Assert.Null(cell.Attribute("Err")));
    }

    [Fact]
    public void DetachedClonePreservesNativeSyntaxWithoutSharingAssignmentState() {
        VisioDocument document = Load("255");
        VisioTextStyle original = document.Pages[0].Shapes[0].TextStyle!;
        VisioTextStyle clone = original.Clone();
        document.Pages[0].Shapes[0].TextStyle = clone;
        AssertCell(Export(document).Descendants(Legacy + "TextBkgnd").First(), "255", "GUARD(255)", "DL");
        clone.BackgroundColor = clone.BackgroundColor;
        AssertCell(Export(document).Descendants(Legacy + "TextBkgnd").First(), "0", null, null);
        document.Pages[0].Shapes[0].TextStyle = original;
        AssertCell(Export(document).Descendants(Legacy + "TextBkgnd").First(), "255", "GUARD(255)", "DL");
    }

    [Fact]
    public void CopiedStylesUseTheirResolvedColorWhenDestinationPaletteConflicts() {
        VisioDocument source = Load("3");
        VisioDocument destination = Load("3", paletteColor: "#654321");
        destination.Pages[0].Shapes[0].ApplyTextStyle(source.Pages[0].Shapes[0].TextStyle!);
        destination.Pages[0].Connectors[0].ApplyTextStyle(source.Pages[0].Connectors[0].TextStyle!);
        foreach (VisioDocument candidate in ReopenedCandidates(destination)) {
            Assert.All(Styles(candidate), style => Assert.Equal(OfficeColor.FromRgb(18, 52, 86), style.BackgroundColor));
            foreach (XElement cell in Export(candidate).Descendants(Legacy + "TextBkgnd"))
                AssertCell(cell, "#123456", "RGB(18,52,86)+1", null);
        }
    }

    [Fact]
    public void SameDocumentDuplicationKeepsPaletteCellsAndRebindsTheirFormulas() {
        XDocument source = Export(Load("3", error: "producer-specific error"));
        XElement shape = source.Descendants(Legacy + "Shape").First();
        shape.Descendants(Legacy + "TextBkgnd").Single().SetAttributeValue("F", "Sheet.1!TextBkgnd");
        shape.Descendants(Legacy + "TextBkgndTrans").Single().SetAttributeValue("F", "Sheet.1!TextBkgndTrans");
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source.ToString()))).Value;
        document.Pages[0].DuplicateShapes(new[] { document.Pages[0].Shapes[0] });
        foreach (VisioDocument candidate in ReopenedCandidates(document)) {
            Assert.Equal(2, candidate.Pages[0].Shapes.Count);
            XElement copied = Export(candidate).Descendants(Legacy + "Shape").Where(shape => shape.Element(Legacy + "Text")?.Value == "Shape").Last();
            string id = (string)copied.Attribute("ID")!;
            XElement background = copied.Descendants(Legacy + "TextBkgnd").Single();
            XElement transparency = copied.Descendants(Legacy + "TextBkgndTrans").Single();
            AssertCell(background, "3", "Sheet." + id + "!TextBkgnd", "DL");
            AssertCell(transparency, "0.2500", "Sheet." + id + "!TextBkgndTrans", "%");
            Assert.Equal("producer-specific error", (string?)background.Attribute("Err"));
            Assert.Equal("producer-specific error", (string?)transparency.Attribute("Err"));
        }
    }

    [Fact]
    public void ImportedMastersUseTheirSourcePaletteBeforeDestinationMerge() {
        const string masterXml = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'>"
            + "<Colors><ColorEntry IX='2' RGB='#123456'/></Colors><Masters><Master ID='0' NameU='Background'>"
            + "<Shapes><Shape ID='1'><XForm><Width>2</Width><Height>1</Height></XForm>"
            + "<TextBlock><TextBkgnd F='GUARD(3)'>3</TextBkgnd></TextBlock><Text>Master</Text>"
            + "</Shape></Shapes></Master></Masters></VisioDocument>";
        VisioDocument source = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(masterXml)), VisioPackageType.Stencil).Value;
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
        try {
            File.WriteAllBytes(path, source.ToBytes());
            VisioDocument destination = Load("3", paletteColor: "#654321");
            var imported = Assert.Single(destination.ImportStencilMastersAndGet(path));
            Assert.Equal(OfficeColor.FromRgb(18, 52, 86), imported.Shape.TextStyle!.BackgroundColor);
            destination.Pages[0].AddShape("imported-background", imported, 5, 5, 2, 1);
            foreach (VisioDocument candidate in ReopenedCandidates(destination)) {
                VisioMaster master = candidate.Masters.Single(master => master.NameU == "Background");
                Assert.Equal(OfficeColor.FromRgb(18, 52, 86), master.Shape.TextStyle!.BackgroundColor);
                XElement background = Export(candidate).Descendants(Legacy + "Master").Single(master => (string?)master.Attribute("NameU") == "Background")
                    .Descendants(Legacy + "TextBkgnd").Single();
                AssertCell(background, "#123456", "RGB(18,52,86)+1", null);
            }
        } finally { File.Delete(path); }
    }

    [Fact]
    public void NewStylesWriteNoFillAndNormalizedTransparencyCaches() {
        VisioDocument document = VisioDocument.Create();
        VisioPage page = document.AddPage("Backgrounds");
        VisioShape first = page.AddRectangle(1, 1, 1, 1, "Shape");
        VisioShape second = page.AddRectangle(3, 1, 1, 1, "Target");
        first.ApplyTextStyle(new VisioTextStyle { BackgroundColor = OfficeColor.Transparent, BackgroundTransparency = 75 });
        page.AddConnector(first, second).ApplyTextStyle(new VisioTextStyle { BackgroundColor = OfficeColor.White, BackgroundTransparency = 25 });
        using var archive = new ZipArchive(new MemoryStream(document.ToBytes()), ZipArchiveMode.Read);
        using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
        XElement[] backgrounds = XDocument.Load(stream).Descendants(Modern + "Cell").Where(cell => (string?)cell.Attribute("N") == "TextBkgnd").ToArray();
        Assert.Equal("0", (string?)backgrounds[0].Attribute("V"));
        Assert.Null(backgrounds[0].Attribute("F"));
        Assert.Equal("#FFFFFF", (string?)backgrounds[1].Attribute("V"));
        Assert.Equal("RGB(255,255,255)+1", (string?)backgrounds[1].Attribute("F"));
        VisioDocument reopened = VisioDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(75, reopened.Pages[0].Shapes[0].TextStyle!.BackgroundTransparency);
        Assert.Equal(25, reopened.Pages[0].Connectors[0].TextStyle!.BackgroundTransparency);
    }

    private static VisioDocument Load(string value, string paletteColor = "#123456", string? error = null) {
        string errorAttribute = error == null ? "" : " Err='" + error + "'";
        string background = "<TextBlock><TextBkgnd Unit='DL' F='GUARD(" + value + ")'" + errorAttribute + ">" + value + "</TextBkgnd>"
            + "<TextBkgndTrans Unit='%' F='GUARD(25%)'" + errorAttribute + ">0.2500</TextBkgndTrans></TextBlock>";
        string xml = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Colors>"
            + "<ColorEntry IX='0' RGB='#000000'/><ColorEntry IX='1' RGB='#FFFFFF'/><ColorEntry IX='2' RGB='" + paletteColor + "'/></Colors>"
            + "<Pages><Page ID='0' Name='Backgrounds'><Shapes><Shape ID='1'><XForm><PinX>1</PinX><PinY>1</PinY><Width>2</Width><Height>1</Height></XForm>"
            + background + "<Text>Shape</Text></Shape><Shape ID='2'><XForm1D><BeginX>1</BeginX><BeginY>1</BeginY><EndX>4</EndX><EndY>1</EndY></XForm1D>"
            + background + "<Text>Connector</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
    }

    private static IEnumerable<VisioTextStyle> Styles(VisioDocument document) =>
        new[] { document.Pages[0].Shapes[0].TextStyle!, document.Pages[0].Connectors[0].TextStyle! };

    private static IEnumerable<VisioDocument> ReopenedCandidates(VisioDocument document) {
        yield return document;
        yield return VisioDocument.Load(new MemoryStream(document.ToBytes()));
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
    }

    private static XDocument Export(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));

    private static void AssertCell(XElement cell, string value, string? formula, string? unit) {
        Assert.Equal(value, cell.Value);
        Assert.Equal(formula, (string?)cell.Attribute("F"));
        Assert.Equal(unit, (string?)cell.Attribute("Unit"));
    }
}

using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioPageSheetNativeUnitsTests {
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly string[] LengthNames = {
        "PageWidth", "PageHeight", "PageScale", "DrawingScale",
        "PageLeftMargin", "PageRightMargin", "PageTopMargin", "PageBottomMargin",
        "BlockSizeX", "BlockSizeY", "AvenueSizeX", "AvenueSizeY",
        "LineToLineX", "LineToLineY", "LineToNodeX", "LineToNodeY"
    };

    [Fact]
    public void NativeVisioMetricFixtureKeepsA4PhysicalSizeAndOriginalDistanceCaches() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "Assets", "VisioTemplates", "DrawingWithLotsOfShapresAndArrows.vsdx");
        XDocument source;
        using (var archive = ZipFile.OpenRead(path)) source = ReadPages(archive);
        VisioDocument document = VisioDocument.Load(path);
        var drawings = document.ToDrawings().Value;
        Assert.Equal(10, document.Pages.Count);
        Assert.Equal(document.Pages.Count, drawings.Count);
        for (int index = 0; index < document.Pages.Count; index++) {
            VisioPage page = document.Pages[index];
            Assert.Equal(11.69291338582677, page.Width, 12);
            Assert.Equal(8.26771653543307, page.Height, 12);
            Assert.Equal(VisioMeasurementUnit.Millimeters, page.DefaultUnit);
            Assert.Equal(1, page.PageScale.Value, 12);
            Assert.Equal(1, page.DrawingScale.Value, 12);
            Assert.Equal(VisioMeasurementUnit.Millimeters, page.PageScale.Unit);
            Assert.Equal(VisioMeasurementUnit.Millimeters, page.DrawingScale.Unit);
            Assert.Equal(0.25, page.LeftMargin, 12);
            Assert.Equal(841.8897637795274, drawings[index].Width, 9);
            Assert.Equal(595.275590551181, drawings[index].Height, 9);
        }
        foreach (VisioDocument candidate in new[] { document, Reopen(document) }) {
            XDocument written = Pages(candidate);
            foreach (XElement original in source.Root!.Elements(Modern + "Page")) {
                XElement saved = written.Root!.Elements(Modern + "Page").Single(page =>
                    (string?)page.Attribute("ID") == (string?)original.Attribute("ID"));
                Assert.Equal(DistanceCells(original.Element(Modern + "PageSheet")!),
                    DistanceCells(saved.Element(Modern + "PageSheet")!));
            }
        }
    }

    [Theory]
    [InlineData("urn:schemas-microsoft-com:office:visio")]
    [InlineData("http://schemas.microsoft.com/visio/2003/core")]
    public void LegacyDistanceCachesUseInternalInchesAndKeepNativeCellStateAcrossPackageAndXml(string ns) {
        VisioDocument document = LoadNativeLengths(ns);
        foreach (VisioDocument candidate in new[] { document, Reopen(document), ReloadLegacy(document) }) {
            VisioPage page = Assert.Single(candidate.Pages);
            Assert.Equal(11.69291338582677, page.Width, 12);
            Assert.Equal(8.26771653543307, page.Height, 12);
            Assert.Equal(0.25, page.LeftMargin, 12);
            Assert.Equal(0.375, page.RightMargin, 12);
            Assert.Equal(0.5, page.TopMargin, 12);
            Assert.Equal(0.625, page.BottomMargin, 12);
            Assert.Equal(0.5, page.LayoutBlockSizeX.GetValueOrDefault(), 12);
            Assert.Equal(0.75, page.LayoutBlockSizeY.GetValueOrDefault(), 12);
            Assert.Equal(0.125, page.LayoutAvenueSizeX.GetValueOrDefault(), 12);
            Assert.Equal(0.375, page.LayoutAvenueSizeY.GetValueOrDefault(), 12);
            Assert.Equal(0.125, page.LineToLineX.GetValueOrDefault(), 12);
            Assert.Equal(0.25, page.LineToLineY.GetValueOrDefault(), 12);
            Assert.Equal(0.375, page.LineToNodeX.GetValueOrDefault(), 12);
            Assert.Equal(0.5, page.LineToNodeY.GetValueOrDefault(), 12);
            Assert.Equal(VisioMeasurementUnit.Millimeters, page.PageScale.Unit);
            Assert.Equal(1, page.PageScale.Value, 12);
            Assert.Equal(VisioMeasurementUnit.Centimeters, page.DrawingScale.Unit);
            Assert.Equal(0.1, page.DrawingScale.Value, 12);
            XElement sheet = LegacySheet(candidate, "Native");
            Assert.Equal("0.2500", LegacyCell(sheet, "PageLeftMargin").Value);
            Assert.Equal("GUARD(6.35mm)", (string?)LegacyCell(sheet, "PageLeftMargin").Attribute("F"));
            Assert.Equal("#VALUE!", (string?)LegacyCell(sheet, "PageLeftMargin").Attribute("Err"));
            Assert.Equal("MM", (string?)LegacyCell(sheet, "PageLeftMargin").Attribute("Unit"));
            Assert.Equal("CM", (string?)LegacyCell(sheet, "LineToLineY").Attribute("Unit"));
            Assert.Equal("MM", (string?)LegacyCell(sheet, "LineToNodeX").Attribute("Unit"));
        }
        Assert.Equal(LegacyDistances(LegacySheet(document, "Native")),
            LegacyDistances(LegacySheet(Reopen(document), "Native")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TypedDistanceEditsWriteInchesAndCopiedNativeFormulasFollowTheirOwnShapes(bool reopenPackage) {
        VisioDocument document = LoadNativeLengths(Legacy.NamespaceName);
        if (reopenPackage) document = Reopen(document);
        VisioPage original = document.Pages[0];
        VisioPage copy = document.DuplicatePage(original, "Copy");
        XElement copiedSheet = LegacySheet(document, "Copy");
        XDocument copiedXml = Export(document);
        string copiedId = copiedXml.Descendants(Legacy + "Page").Single(page => (string?)page.Attribute("Name") == "Copy")
            .Element(Legacy + "Shapes")!.Element(Legacy + "Shape")!.Attribute("ID")!.Value;
        Assert.Equal("GUARD(Sheet." + copiedId + "!Width)", (string?)LegacyCell(copiedSheet, "PageWidth").Attribute("F"));
        Assert.Equal("GUARD(Sheet.7!Width)", (string?)LegacyCell(LegacySheet(document, "Native"), "PageWidth").Attribute("F"));
        var sourceCells = LegacySheet(document, "Native");
        LegacyCell(sourceCells, "PageWidth").SetAttributeValue("F", "GUARD(Sheet." + copiedId + "!Width)");
        Assert.Equal(LegacyDistances(sourceCells), LegacyDistances(copiedSheet));

        original.WidthCentimeters = 12.7;
        original.DefaultUnit = VisioMeasurementUnit.Centimeters;
        original.SetMargins(1.27, VisioMeasurementUnit.Centimeters);
        original.SetLayoutGridSizing(25.4, 50.8, 12.7, 6.35, VisioMeasurementUnit.Millimeters);
        original.SetConnectorSpacing(2.54, 5.08, 1.27, 0.635, VisioMeasurementUnit.Centimeters);
        original.PageScale = new VisioScaleSetting(2.54, VisioMeasurementUnit.Centimeters);
        original.DrawingScale = new VisioScaleSetting(50.8, VisioMeasurementUnit.Millimeters);
        foreach (VisioDocument candidate in new[] { document, Reopen(document), ReloadLegacy(document) }) {
            XElement sheet = Pages(candidate).Root!.Elements(Modern + "Page")
                .Single(page => (string?)page.Attribute("Name") == "Native").Element(Modern + "PageSheet")!;
            AssertCache(sheet, "PageWidth", 5, "CM");
            AssertCache(sheet, "PageHeight", 8.26771653543307, "CM");
            AssertCache(sheet, "PageScale", 1, "CM");
            AssertCache(sheet, "DrawingScale", 2, "MM");
            foreach (string name in new[] { "PageLeftMargin", "PageRightMargin", "PageTopMargin", "PageBottomMargin" })
                AssertCache(sheet, name, 0.5, "CM");
            AssertCache(sheet, "BlockSizeX", 1, "MM");
            AssertCache(sheet, "BlockSizeY", 2, "MM");
            AssertCache(sheet, "AvenueSizeX", 0.5, "MM");
            AssertCache(sheet, "AvenueSizeY", 0.25, "MM");
            AssertCache(sheet, "LineToLineX", 1, "CM");
            AssertCache(sheet, "LineToLineY", 2, "CM");
            AssertCache(sheet, "LineToNodeX", 0.5, "CM");
            AssertCache(sheet, "LineToNodeY", 0.25, "CM");
            Assert.Null(LegacyCell(LegacySheet(candidate, "Native"), "PageLeftMargin").Attribute("Err"));
            Assert.Equal(11.69291338582677, candidate.Pages.Single(page => page.Name == "Copy").Width, 12);
            Assert.Equal(LegacyDistances(copiedSheet), LegacyDistances(LegacySheet(candidate, "Copy")));
        }
        Assert.Equal(11.69291338582677, copy.Width, 12);
    }

    private static VisioDocument LoadNativeLengths(string ns) {
        string source = "<VisioDocument xmlns='" + ns + "'><Pages><Page ID='0' Name='Native'><PageSheet><PageProps>"
            + "<PageWidth Unit='MM' F='GUARD(Sheet.7!Width)'>11.69291338582677</PageWidth><PageHeight Unit='MM'>8.26771653543307</PageHeight>"
            + "<PageScale Unit='MM'>0.03937007874015748</PageScale><DrawingScale Unit='CM'>0.03937007874015748</DrawingScale></PageProps><PrintProps>"
            + "<PageLeftMargin Unit='MM' F='GUARD(6.35mm)' Err='#VALUE!'>0.2500</PageLeftMargin><PageRightMargin Unit='CM'>0.375</PageRightMargin>"
            + "<PageTopMargin Unit='IN'>0.5</PageTopMargin><PageBottomMargin Unit='MM'>0.625</PageBottomMargin></PrintProps><PageLayout>"
            + "<BlockSizeX Unit='MM'>0.5</BlockSizeX><BlockSizeY Unit='CM'>0.75</BlockSizeY><AvenueSizeX Unit='IN'>0.125</AvenueSizeX><AvenueSizeY Unit='MM'>0.375</AvenueSizeY>"
            + "<LineToLineX Unit='IN'>0.125</LineToLineX><LineToLineY Unit='CM'>0.25</LineToLineY><LineToNodeX Unit='MM'>0.375</LineToNodeX><LineToNodeY Unit='IN'>0.5</LineToNodeY>"
            + "</PageLayout></PageSheet><Shapes><Shape ID='7'><XForm><PinX>1</PinX><PinY>1</PinY><Width>2</Width><Height>1</Height></XForm></Shape></Shapes></Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
    }

    private static void AssertCache(XElement sheet, string name, double inches, string unit) {
        XElement cell = sheet.Elements(Modern + "Cell").Single(value => (string?)value.Attribute("N") == name);
        Assert.Equal(inches, double.Parse((string)cell.Attribute("V")!, CultureInfo.InvariantCulture), 12);
        Assert.Equal(unit, (string?)cell.Attribute("U"));
        Assert.Null(cell.Attribute("F"));
        Assert.Null(cell.Attribute("E"));
    }

    private static string[] DistanceCells(XElement sheet) => sheet.Elements(Modern + "Cell")
        .Where(cell => LengthNames.Contains((string?)cell.Attribute("N"))).Select(cell => cell.ToString(SaveOptions.DisableFormatting)).OrderBy(value => value, StringComparer.Ordinal).ToArray();
    private static string[] LegacyDistances(XElement sheet) => sheet.Descendants()
        .Where(cell => LengthNames.Contains(cell.Name.LocalName)).Select(cell => cell.ToString(SaveOptions.DisableFormatting)).OrderBy(value => value, StringComparer.Ordinal).ToArray();
    private static XElement LegacyCell(XElement sheet, string name) => sheet.Descendants(Legacy + name).Single();
    private static XElement LegacySheet(VisioDocument document, string name) => Export(document).Descendants(Legacy + "Page")
        .Single(page => (string?)page.Attribute("Name") == name).Element(Legacy + "PageSheet")!;
    private static VisioDocument Reopen(VisioDocument document) => VisioDocument.Load(new MemoryStream(document.ToBytes()));
    private static VisioDocument ReloadLegacy(VisioDocument document) => VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
    private static XDocument Export(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
    private static XDocument Pages(VisioDocument document) {
        using var archive = new ZipArchive(new MemoryStream(document.ToBytes()), ZipArchiveMode.Read);
        return ReadPages(archive);
    }
    private static XDocument ReadPages(ZipArchive archive) {
        using Stream stream = archive.GetEntry("visio/pages/pages.xml")!.Open();
        return XDocument.Load(stream);
    }
}

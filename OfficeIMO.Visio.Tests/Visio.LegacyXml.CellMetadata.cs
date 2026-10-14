using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioLegacyCellMetadataTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";

    [Fact]
    public void IndependentNxbreNullStringCellsSurviveXmlAndPackageReopen() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "nxbre-ie3.vsx");
        var original = XDocument.Load(path);
        var document = VisioDocument.LoadLegacyXml(path).Value;
        var expected = NativeNullCells(original);
        Assert.NotEmpty(expected);
        foreach (var candidate in new[] { document, ReopenPackage(document) }) {
            var exported = Export(candidate);
            Assert.Equal(expected, NativeNullCells(exported));
            var xmlReopened = VisioDocument.LoadLegacyXml(new MemoryStream(candidate.ToLegacyXmlResult().Value), VisioPackageType.Stencil).Value;
            Assert.Equal(expected, NativeNullCells(Export(xmlReopened)));
        }
    }

    [Theory]
    [InlineData("http://schemas.microsoft.com/visio/2003/core")]
    [InlineData("urn:schemas-microsoft-com:office:visio")]
    public void ShapeAndMasterCellsKeepLiteralNullFormulaUnitAndErrorMetadata(string ns) {
        var document = Load(ns);
        foreach (var candidate in new[] { document, ReopenPackage(document) }) {
            var xml = Export(candidate);
            Assert.Equal(2, xml.Descendants(Legacy + "Shape").Count());
            foreach (var shape in xml.Descendants(Legacy + "Shape")) {
                var para = shape.Element(Legacy + "Para")!;
                AssertCell(para.Element(Legacy + "BulletStr")!, "", "null", "Inh", "STR", null);
                AssertCell(shape.Element(Legacy + "Misc")!.Element(Legacy + "Comment")!, "", "null", null, null, null);
                var user = shape.Elements(Legacy + "User").Single(row => (string?)row.Attribute("NameU") == "Nullable");
                AssertCell(user.Element(Legacy + "Value")!, "", "null", "GUARD(\"\")", "STR", "#VALUE!");
                AssertCell(user.Element(Legacy + "Prompt")!, "", "null", "Inh", "STR", null);
                var scratch = shape.Element(Legacy + "Scratch")!;
                AssertCell(scratch.Element(Legacy + "X")!, "1.2500", null, "1/0", "IN", "#DIV/0!");
            }
            AssertModernCellMetadata(candidate.ToBytes());
        }
    }

    [Theory]
    [InlineData("value")]
    [InlineData("formula")]
    [InlineData("unit")]
    public void ChangedSourceSectionDoesNotRestoreItsNativeNullMarker(string edit) {
        var document = Load(Legacy.NamespaceName);
        foreach (var shape in new[] { document.Pages[0].Shapes[0], document.Masters.First().Shape }) {
            var paragraph = shape.GetShapeSheetSections().Single(section => section.Name == "Paragraph");
            paragraph.Rows[0].SetCell("BulletStr", edit == "value" ? "bullet" : "",
                edit == "formula" ? "\"bullet\"" : "Inh", edit == "unit" ? "IN" : "STR");
            shape.SetShapeSheetSection(paragraph);
        }
        foreach (var candidate in new[] { document, ReopenPackage(document) }) {
            var xml = Export(candidate);
            foreach (var shape in xml.Descendants(Legacy + "Shape")) {
                var bullet = shape.Element(Legacy + "Para")!.Element(Legacy + "BulletStr")!;
                Assert.Null(bullet.Attribute("V"));
                Assert.Equal(edit == "value" ? "bullet" : "", bullet.Value);
                Assert.Equal(edit == "formula" ? "\"bullet\"" : "Inh", (string?)bullet.Attribute("F"));
                Assert.Equal(edit == "unit" ? "IN" : "STR", (string?)bullet.Attribute("Unit"));
                Assert.Equal("null", (string?)shape.Element(Legacy + "Misc")!.Element(Legacy + "Comment")!.Attribute("V"));
            }
        }
    }

    [Theory]
    [InlineData("value")]
    [InlineData("formula")]
    [InlineData("unit")]
    public void ChangedModeledCellDoesNotRestoreItsNativeNullMarker(string edit) {
        var document = Load(Legacy.NamespaceName);
        foreach (var shape in new[] { document.Pages[0].Shapes[0], document.Masters.First().Shape }) {
            var cell = shape.FindUserCell("Nullable")!;
            if (edit == "value") shape.SetUserCell("Nullable", "edited", "STR");
            else if (edit == "formula") cell.Formula = "\"changed\"";
            else cell.Unit = "IN";
        }
        foreach (var candidate in new[] { document, ReopenPackage(document) }) {
            foreach (var value in Export(candidate).Descendants(Legacy + "User")
                         .Where(row => (string?)row.Attribute("NameU") == "Nullable").Elements(Legacy + "Value")) {
                Assert.Null(value.Attribute("V"));
                if (edit == "value") Assert.Equal("edited", value.Value);
                if (edit == "formula") Assert.Equal("\"changed\"", (string?)value.Attribute("F"));
                if (edit == "unit") Assert.Equal("IN", (string?)value.Attribute("Unit"));
                if (edit != "unit") Assert.Null(value.Attribute("Err"));
            }
        }
    }

    [Fact]
    public void NoncanonicalLegacyErrorIsReportedAndPreservedWithoutInvalidModernErrorToken() {
        const string source = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Pages><Page ID='0'><Shapes><Shape ID='1'>"
            + "<XForm><Width>2</Width><Height>1</Height></XForm><Scratch IX='0'><X Unit='IN' F='CustomFormula' Err='producer-specific error'>1.2500</X>"
            + "</Scratch></Shape></Shapes></Page></Pages></VisioDocument>";
        var imported = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source)));
        Assert.Contains(imported.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "VDX_CELL_ERROR" && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        var document = imported.Value;
        foreach (var candidate in new[] { document, ReopenPackage(document) }) {
            AssertCell(Export(candidate).Descendants(Legacy + "Scratch").Single().Element(Legacy + "X")!,
                "1.2500", null, "CustomFormula", "IN", "producer-specific error");
            using var archive = new ZipArchive(new MemoryStream(candidate.ToBytes()), ZipArchiveMode.Read);
            using var page = archive.GetEntry("visio/pages/page1.xml")!.Open();
            var cell = XDocument.Load(page).Descendants(Modern + "Cell").Single(element => (string?)element.Attribute("F") == "CustomFormula");
            Assert.Null(cell.Attribute("E"));
            Assert.Null(cell.Attribute("Err"));
            Assert.Equal("1.2500", (string?)cell.Attribute("V"));
        }
        var shape = document.Pages[0].Shapes[0];
        var scratch = shape.GetShapeSheetSections().Single(section => section.Name == "Scratch");
        scratch.Rows[0].SetCell("X", "2", unit: "IN");
        shape.SetShapeSheetSection(scratch);
        Assert.Null(Export(ReopenPackage(document)).Descendants(Legacy + "Scratch").Single().Element(Legacy + "X")!.Attribute("Err"));
    }

    private static VisioDocument Load(string ns) {
        const string shape = "<Shape ID='1'><XForm><Width>2</Width><Height>1</Height></XForm>"
            + "<Misc><Comment V='null'/></Misc><Para IX='0'><BulletStr V='null' Unit='STR' F='Inh'/></Para>"
            + "<User ID='2' NameU='Nullable'><Value V='null' Unit='STR' F='GUARD(&quot;&quot;)' Err='#VALUE!'/><Prompt V='null' Unit='STR' F='Inh'/></User>"
            + "<Scratch IX='0'><X Unit='IN' F='1/0' Err='#DIV/0!'>1.2500</X></Scratch><Text>label</Text></Shape>";
        string source = "<VisioDocument xmlns='" + ns + "'><Masters><Master ID='0' NameU='Null master'><Shapes>" + shape
            + "</Shapes></Master></Masters><Pages><Page ID='0'><Shapes>" + shape + "</Shapes></Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source)), VisioPackageType.Template).Value;
    }

    private static VisioDocument ReopenPackage(VisioDocument document) => VisioDocument.Load(new MemoryStream(document.ToBytes()));
    private static XDocument Export(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));

    private static string[] NativeNullCells(XDocument xml) => xml.Descendants().Where(cell => cell.Attribute("V") != null)
        .Select(cell => string.Join("/", cell.AncestorsAndSelf().Reverse().Select(element => element.Name.LocalName
            + ":" + (string?)element.Attribute("NameU") + ":" + (string?)element.Attribute("ID") + ":" + (string?)element.Attribute("IX")))
            + "=" + cell.Value + ";V=" + (string?)cell.Attribute("V") + ";F=" + (string?)cell.Attribute("F") + ";U=" + (string?)cell.Attribute("Unit"))
        .OrderBy(value => value, StringComparer.Ordinal).ToArray();

    private static void AssertCell(XElement cell, string value, string? marker, string? formula, string? unit, string? error) {
        Assert.Equal(value, cell.Value);
        Assert.Equal(marker, (string?)cell.Attribute("V"));
        Assert.Equal(formula, (string?)cell.Attribute("F"));
        Assert.Equal(unit, (string?)cell.Attribute("Unit"));
        Assert.Equal(error, (string?)cell.Attribute("Err"));
        Assert.Null(cell.Attribute("E"));
    }

    private static void AssertModernCellMetadata(byte[] bytes) {
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        foreach (var entry in archive.Entries.Where(entry => entry.FullName.StartsWith("visio/", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))) {
            using var stream = entry.Open();
            var xml = XDocument.Load(stream);
            foreach (var cell in xml.Descendants(Modern + "Cell")) {
                Assert.All(cell.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration), attribute => {
                    Assert.Equal(XNamespace.None, attribute.Name.Namespace);
                    Assert.Contains(attribute.Name.LocalName, new[] { "N", "U", "E", "F", "V" });
                });
                Assert.Null(cell.Attribute("Err"));
                if ((string?)cell.Attribute("N") == "X" && (string?)cell.Attribute("F") == "1/0") {
                    Assert.Equal("1.2500", (string?)cell.Attribute("V"));
                    Assert.Equal("#DIV/0!", (string?)cell.Attribute("E"));
                }
            }
        }
    }
}

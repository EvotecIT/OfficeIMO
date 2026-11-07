using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordMlReconstructionTests {
    private const string WordNamespace = "http://schemas.microsoft.com/office/word/2003/wordml";

    [Theory]
    [InlineData("hdr", "<w:sectPr/>")]
    [InlineData("ftr", "<w:sectPr/>")]
    [InlineData("hdr", "<w:p><w:pPr><w:sectPr/></w:pPr></w:p>")]
    [InlineData("ftr", "<w:p><w:pPr><w:sectPr/></w:pPr></w:p>")]
    [InlineData("hdr", "<wx:sect><wx:sub-section><w:sectPr/></wx:sub-section></wx:sect>")]
    [InlineData("ftr", "<wx:sect><wx:sub-section><w:p><w:pPr><w:sectPr/></w:pPr></w:p></wx:sub-section></wx:sect>")]
    public void SectionPropertiesOutsideTheMainStoryAreOmittedAndReported(string story, string block) {
        string xml = Wrap("<w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr>" +
            "<w:" + story + " w:type=\"odd\"><w:p><w:r><w:t>Story</w:t></w:r></w:p>" + block +
            "</w:" + story + "></w:sectPr></w:body>");
        using var input = Input(xml);
        using WordDocument document = WordDocument.Load(input, null, out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Location?.EndsWith("}sectPr", StringComparison.Ordinal) == true &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
        using var saved = new MemoryStream(document.ToBytes());
        using WordprocessingDocument package = WordprocessingDocument.Open(saved, false);
        OpenXmlPartRootElement root = story == "hdr"
            ? Assert.Single(package.MainDocumentPart!.HeaderParts).Header
            : Assert.Single(package.MainDocumentPart!.FooterParts).Footer;
        Assert.Contains("Story", root.InnerText);
        Assert.Empty(root.Descendants<SectionProperties>());
        Assert.Single(package.MainDocumentPart.Document.Body!.Elements<SectionProperties>());
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package));
        using var strict = Input(xml);
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(strict, new WordLoadOptions { FailOnLoss = true }));
    }

    [Theory]
    [InlineData("styles")]
    [InlineData("fonts")]
    [InlineData("docPr")]
    [InlineData("metadata")]
    [InlineData("defaultFonts")]
    [InlineData("hdr")]
    [InlineData("ftr")]
    public void DuplicateSingletonDeclarationsFailWithSourceAdmissionError(string kind) {
        string element = kind switch {
            "styles" => "<w:styles/>",
            "fonts" => "<w:fonts><w:font w:name=\"Arial\"/></w:fonts>",
            "docPr" => "<w:docPr><w:view w:val=\"print\"/></w:docPr>",
            "metadata" => "<o:DocumentProperties><o:Title>Title</o:Title></o:DocumentProperties>",
            "defaultFonts" => "<w:defaultFonts w:ascii=\"Arial\"/>",
            _ => "<w:" + kind + " w:type=\"odd\"><w:p/></w:" + kind + ">"
        };
        string content = kind is "hdr" or "ftr"
            ? "<w:body><w:sectPr>" + element + element + "</w:sectPr></w:body>"
            : (kind == "defaultFonts" ? "<w:fonts>" + element + element + "</w:fonts>" : element + element) + "<w:body><w:p/></w:body>";
        using var input = Input(Wrap(content));
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(input));
        Assert.Contains("duplicate", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ReportedVmlLockLossIsAbsentFromTheSavedPicture() {
        string xml = Wrap("""
            <w:body><w:p><w:r><w:pict><v:shape id="LockedPicture" style="width:12pt;height:6pt">
              <o:lock xmlns:producer="urn:producer:metadata" v:ext="edit" aspectratio="t" unsupported="secret">
                <v:path v="m0,0l1,1"/>unexpected lock text
              </o:lock>
            </v:shape></w:pict></w:r></w:p></w:body>
            """);
        using var input = Input(xml);
        using WordDocument document = WordDocument.Load(input, null, out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Location?.EndsWith("}unsupported", StringComparison.Ordinal) == true ||
            diagnostic.Location == "unsupported");
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Message.Contains("lock", StringComparison.Ordinal) &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
        using var saved = new MemoryStream(document.ToBytes());
        using WordprocessingDocument package = WordprocessingDocument.Open(saved, false);
        XElement output = XElement.Parse(package.MainDocumentPart!.Document.OuterXml);
        XElement locked = Assert.Single(output.Descendants(XName.Get("lock", "urn:schemas-microsoft-com:office:office")));
        Assert.Equal("t", (string?)locked.Attribute("aspectratio"));
        Assert.Equal("edit", (string?)locked.Attribute(XName.Get("ext", "urn:schemas-microsoft-com:vml")));
        Assert.Equal("urn:producer:metadata", locked.GetNamespaceOfPrefix("producer")?.NamespaceName);
        Assert.Null(locked.Attribute("unsupported"));
        Assert.Empty(locked.Elements());
        Assert.DoesNotContain("unexpected lock text", locked.Value);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package));
        using var strict = Input(xml);
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(strict, new WordLoadOptions { FailOnLoss = true }));
    }

    [Theory]
    [InlineData("<v:formulas><v:shape id=\"Nested\"/></v:formulas>", "shape")]
    [InlineData("<v:formulas><v:f eqn=\"val 1\"><v:path v=\"m0,0\"/></v:f></v:formulas>", "path")]
    [InlineData("<v:fill><o:lock aspectratio=\"t\"/></v:fill>", "lock")]
    [InlineData("<v:formulas><v:imagedata src=\"wordml://pixel.png\"/></v:formulas>", "imagedata")]
    [InlineData("<v:formulas>unexpected text<v:f eqn=\"val 1\"/></v:formulas>", "formulas")]
    public void VmlChildContextLossIsReportedBeforeResourceCreationAndEmission(string content, string omittedName) {
        string xml = Wrap("<w:body><w:p><w:r><w:pict><w:binData w:name=\"wordml://pixel.png\">" +
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=" +
            "</w:binData><v:shape id=\"Picture1\" style=\"width:12pt;height:6pt\">" + content +
            "</v:shape></w:pict></w:r></w:p></w:body>");
        using var input = Input(xml);
        using WordDocument document = WordDocument.Load(input, null, out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Location?.EndsWith("}" + omittedName, StringComparison.Ordinal) == true &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
        using var saved = new MemoryStream(document.ToBytes());
        using WordprocessingDocument package = WordprocessingDocument.Open(saved, false);
        Assert.Empty(package.MainDocumentPart!.ImageParts);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package));
        using var strict = Input(xml);
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(strict, new WordLoadOptions { FailOnLoss = true }));
    }

    [Theory]
    [InlineData("Tm90IGFuIGltYWdl")]
    [InlineData("")]
    [InlineData("iVBORw0KGgo=")]
    public void UnrecognizedEmbeddedImagePayloadsFailWithSourceAdmissionError(string payload) {
        string xml = Wrap("<w:body><w:p><w:r><w:pict><w:binData w:name=\"wordml://pixel.png\">" + payload +
            "</w:binData><v:shape id=\"Picture1\" style=\"width:12pt;height:6pt\">" +
            "<v:imagedata src=\"wordml://pixel.png\"/></v:shape></w:pict></w:r></w:p></w:body>");
        using var input = Input(xml);
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(input));
        Assert.Contains("WordML 2003 contains an unsupported image payload", error.Message, StringComparison.Ordinal);
        Assert.IsType<NotSupportedException>(error.InnerException);
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData("xml")]
    [InlineData("part")]
    [InlineData("aggregate")]
    [InlineData("offsets")]
    [InlineData("explicit")]
    [InlineData("offset-overflow")]
    public void GridGeometryIsBoundedDuringReconstructionBeforePackageInspection(string boundary) {
        string table = Table("64");
        var policy = new OfficePackageSecurityOptions();
        switch (boundary) {
            case "xml": policy.MaxXmlCharactersInPart = 639; break;
            case "part": policy.MaxPartUncompressedBytes = 639; break;
            case "aggregate": table = Table("32") + Table("32"); policy.MaxTotalUncompressedBytes = 639; break;
            case "offsets": table = Table(null, "32", "32"); policy.MaxPartUncompressedBytes = 649; break;
            case "explicit":
                table = Table(null).Replace("<w:tr>", "<w:tblGrid>" + string.Concat(Enumerable.Repeat("<w:gridCol/>", 64)) + "</w:tblGrid><w:tr>");
                policy.MaxXmlCharactersInPart = 639;
                break;
            default: table = Table(null, "2147483647", "2147483647"); break;
        }
        using var output = new MemoryStream();
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordMl2003Projection.WritePackage(
            XDocument.Parse(Wrap("<w:body>" + table + "</w:body>")).Root!, output, policy, default));
        Assert.Contains("table grid", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void InferredGridIncludesRowOffsetsAndRetainsStructuredCells() {
        string xml = Wrap("<w:body>" + Table("2", "1", "2") + "</w:body>");
        using var input = Input(xml);
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions { FailOnLoss = true });
        using var saved = new MemoryStream(document.ToBytes());
        using WordprocessingDocument package = WordprocessingDocument.Open(saved, false);
        Table table = Assert.Single(package.MainDocumentPart!.Document.Body!.Elements<Table>());
        Assert.Equal(5, table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());
        TableRow row = Assert.Single(table.Elements<TableRow>());
        Assert.Equal(1, row.TableRowProperties!.GetFirstChild<GridBefore>()!.Val!.Value);
        Assert.Equal(2, row.TableRowProperties.GetFirstChild<GridAfter>()!.Val!.Value);
        Assert.Equal(2, Assert.Single(row.Elements<TableCell>()).TableCellProperties!.GridSpan!.Val!.Value);
        Assert.Equal("Cell", row.InnerText);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package));
    }

    private static string Table(string? span, string? before = null, string? after = null) =>
        "<w:tbl><w:tr>" + (before == null && after == null ? string.Empty : "<w:trPr>" +
        (before == null ? string.Empty : "<w:gridBefore w:val=\"" + before + "\"/>") +
        (after == null ? string.Empty : "<w:gridAfter w:val=\"" + after + "\"/>") + "</w:trPr>") +
        "<w:tc>" + (span == null ? string.Empty : "<w:tcPr><w:gridSpan w:val=\"" + span + "\"/></w:tcPr>") +
        "<w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";

    private static string Wrap(string content) => "<w:wordDocument xmlns:w=\"" + WordNamespace +
        "\" xmlns:wx=\"http://schemas.microsoft.com/office/word/2003/auxHint\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\">" + content + "</w:wordDocument>";

    private static MemoryStream Input(string xml) => new(Encoding.UTF8.GetBytes(xml));
}

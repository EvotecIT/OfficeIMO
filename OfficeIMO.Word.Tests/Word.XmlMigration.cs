using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordMl2003MigrationTests {
    [Fact]
    public void LegacyFormattingEnumsMigrateToValidNativePropertiesUnderStrictLoad() {
        const string xml = """
            <w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml"><w:body>
              <w:p><w:pPr><w:pBdr><w:top w:val="thin-thick-small-gap" w:sz="8"/></w:pBdr>
                <w:shd w:val="pct-5" w:fill="DDEEFF"/><w:tabs><w:tab w:val="left" w:leader="middle-dot" w:pos="720"/></w:tabs>
                <w:spacing w:line="300" w:line-rule="at-least"/><w:jc w:val="medium-kashida"/>
                <w:textDirection w:val="tb-rl"/></w:pPr>
                <w:r><w:rPr><w:u w:val="dash-long-heavy"/><w:effect w:val="ants-black"/><w:em w:val="under-dot"/></w:rPr><w:t>Qualified formatting</w:t></w:r>
              </w:p><w:sectPr><w:type w:val="even-page"/><w:pgNumType w:fmt="upper-roman" w:chap-sep="em-dash"/>
                <w:docGrid w:type="lines-and-chars" w:line-pitch="360"/></w:sectPr>
            </w:body></w:wordDocument>
            """;
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions { FailOnLoss = true }, out OfficeConversionReport report);
        Assert.False(report.HasLoss);
        using var bytes = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(saved));
        ParagraphProperties paragraph = saved.MainDocumentPart.Document.Body.Elements<Paragraph>().Single().ParagraphProperties;
        Assert.Equal(LineSpacingRuleValues.AtLeast, paragraph.SpacingBetweenLines.LineRule.Value);
        Assert.Equal(TextDirectionValues.TopToBottomRightToLeft, paragraph.TextDirection.Val.Value);
        Assert.Equal(ShadingPatternValues.Percent5, paragraph.Shading.Val.Value);
        Assert.Equal(BorderValues.ThinThickSmallGap, paragraph.ParagraphBorders.TopBorder.Val.Value);
        Assert.Equal(TabStopLeaderCharValues.MiddleDot, paragraph.Tabs.GetFirstChild<TabStop>().Leader.Value);
        Assert.Equal(JustificationValues.MediumKashida, paragraph.Justification.Val.Value);
        RunProperties run = saved.MainDocumentPart.Document.Descendants<Run>().Single().RunProperties;
        Assert.Equal(UnderlineValues.DashLongHeavy, run.Underline.Val.Value);
        Assert.Equal(TextEffectValues.AntsBlack, run.TextEffect.Val.Value);
        Assert.Equal(EmphasisMarkValues.UnderDot, run.Emphasis.Val.Value);
        SectionProperties section = saved.MainDocumentPart.Document.Body.GetFirstChild<SectionProperties>();
        Assert.Equal(SectionMarkValues.EvenPage, section.GetFirstChild<SectionType>().Val.Value);
        Assert.Equal(NumberFormatValues.UpperRoman, section.GetFirstChild<PageNumberType>().Format.Value);
        Assert.Equal(ChapterSeparatorValues.EmDash, section.GetFirstChild<PageNumberType>().ChapterSeparator.Value);
        Assert.Equal(DocGridValues.LinesAndChars, section.GetFirstChild<DocGrid>().Type.Value);
    }

    [Theory]
    [InlineData("<w:shd w:val=\"pct-6\"/>")]
    [InlineData("<w:jc w:val=\"unqualified-alignment\"/>")]
    public void UnqualifiedEnumValueIsOmittedAndRemainsAStrictLoss(string property) {
        string xml = "<w:wordDocument xmlns:w=\"http://schemas.microsoft.com/office/word/2003/wordml\"><w:body><w:p><w:pPr>" +
            property + "</w:pPr><w:r><w:t>Retained text</w:t></w:r></w:p></w:body></w:wordDocument>";
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        using WordDocument document = WordDocument.Load(input, null, out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission && diagnostic.Message.Contains("enum value"));
        using var bytes = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(saved));
        Assert.Equal("Retained text", saved.MainDocumentPart.Document.Body.InnerText);
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(input, new WordLoadOptions { FailOnLoss = true, ReportUnsupportedContent = false }));
    }

    [Theory]
    [InlineData(1, false)]
    [InlineData(2, true)]
    public void ReconstructionChargesEveryStoryImagePartAndRetainsPerStoryReuse(int imageBudget, bool succeeds) {
        XElement source = XDocument.Parse(RichDocument()).Root;
        source.Descendants(XName.Get("tbl", "http://schemas.microsoft.com/office/word/2003/wordml")).Remove();
        using var output = new MemoryStream();
        var policy = new OfficePackageSecurityOptions { MaxTotalUncompressedBytes = PixelBytes.Length * imageBudget };
        if (!succeeds) {
            InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordMl2003Projection.WritePackage(source, output, policy, default));
            Assert.Contains("written image resource limit", error.Message);
            return;
        }
        WordMl2003Projection.WritePackage(source, output, policy, default);
        output.Position = 0;
        using WordprocessingDocument package = WordprocessingDocument.Open(output, false);
        Assert.Equal(PixelBytes, ReadBytes(Assert.Single(package.MainDocumentPart.ImageParts)));
        Assert.Equal(PixelBytes, ReadBytes(Assert.Single(Assert.Single(package.MainDocumentPart.HeaderParts).ImageParts)));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ReconstructionReservesNonImagePartsAndTheirRelationshipsBeforeCreation(bool parts) {
        const string xml = """
            <w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml">
            <w:styles><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>
            <w:body><w:p/></w:body></w:wordDocument>
            """;
        var policy = new OfficePackageSecurityOptions();
        if (parts) policy.MaxPartCount = 4;
        else policy.MaxRelationships = 1;
        using var output = new MemoryStream();
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordMl2003Projection.WritePackage(XDocument.Parse(xml).Root, output, policy, default));
        Assert.Contains("relationship or part limit", error.Message);
        if (parts) policy.MaxPartCount++;
        else policy.MaxRelationships++;
        using var accepted = new MemoryStream();
        WordMl2003Projection.WritePackage(XDocument.Parse(xml).Root, accepted, policy, default);
        accepted.Position = 0;
        using WordprocessingDocument package = WordprocessingDocument.Open(accepted, false);
        Assert.Equal("Normal", Assert.Single(package.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>()).StyleId.Value);
    }

    [Theory]
    [InlineData(true, "decoded image resource limit")]
    [InlineData(false, "relationship or part limit")]
    public void WordMlResourceLimitsRejectBeforeCompletingPackageReconstruction(bool imageLimit, string message) {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(RichDocument()));
        var policy = new OfficePackageSecurityOptions();
        if (imageLimit) policy.MaxPartUncompressedBytes = PixelBytes.Length - 1;
        else policy.MaxRelationships = 1;
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => WordDocument.Load(input,
            new WordLoadOptions { PackageSecurity = policy }));
        Assert.Contains(message, error.Message);
    }

    [Fact]
    public void FlatOpcExportPreservesNativeAuxiliaryPartsAndRelationshipGraph() {
        XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "WordXmlMigration", "word-rich-native-flat-opc.xml");
        XDocument source = XDocument.Load(path);
        using WordDocument document = WordDocument.Load(path, new WordLoadOptions { FailOnLoss = true });
        using var stream = new MemoryStream(document.ExportFlatOpc());
        XDocument exported = XDocument.Load(stream);
        var expected = source.Root.Elements(pkg + "part").ToDictionary(part => part.Attribute(pkg + "name")!.Value);
        var actual = exported.Root.Elements(pkg + "part").ToDictionary(part => part.Attribute(pkg + "name")!.Value);
        foreach (var part in expected) {
            Assert.True(actual.ContainsKey(part.Key), "Missing native package part: " + part.Key);
            Assert.Equal((string)part.Value.Attribute(pkg + "contentType"), (string)actual[part.Key].Attribute(pkg + "contentType"));
            if (part.Value.Element(pkg + "binaryData") is { } binary)
                Assert.Equal(Convert.FromBase64String(binary.Value), Convert.FromBase64String(actual[part.Key].Element(pkg + "binaryData").Value));
            if (part.Key.EndsWith(".rels", StringComparison.Ordinal)) {
                int directory = part.Key.LastIndexOf("/_rels/", StringComparison.Ordinal);
                string owner = part.Key == "/_rels/.rels" ? "/" : part.Key.Substring(0, directory + 1) +
                    part.Key.Substring(directory + 7, part.Key.Length - directory - 7 - ".rels".Length);
                string[] Graph(XElement root) => root.Descendants(rel + "Relationship")
                    .Select(edge => {
                        string mode = (string)edge.Attribute("TargetMode") ?? "Internal";
                        string target = (string)edge.Attribute("Target");
                        if (mode == "Internal") target = new Uri(new Uri("https://package.invalid" + owner), target).AbsolutePath;
                        return string.Join("|", (string)edge.Attribute("Id"), (string)edge.Attribute("Type"), target, mode);
                    })
                    .OrderBy(edge => edge, StringComparer.Ordinal).ToArray();
                Assert.Equal(Graph(part.Value), Graph(actual[part.Key]));
            }
        }
        foreach (string name in new[] { "/word/footnotes.xml", "/word/endnotes.xml", "/word/theme/theme1.xml", "/word/webSettings.xml" }) {
            XElement Normalize(XElement part) {
                var root = new XElement(part.Element(pkg + "xmlData").Elements().Single());
                foreach (XElement element in root.DescendantsAndSelf()) {
                    XAttribute[] attributes = element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
                        .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal).Select(attribute => new XAttribute(attribute)).ToArray();
                    element.RemoveAttributes();
                    element.Add(attributes);
                }
                return root;
            }
            Assert.True(XNode.DeepEquals(Normalize(expected[name]), Normalize(actual[name])), "Changed native auxiliary part: " + name);
        }
    }

    [Fact]
    public void RichWordMlMigrationPreservesStylesTableGeometryLinksAndStoryResources() {
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(RichDocument()));
        using WordDocument document = WordDocument.Load(input, new WordLoadOptions { FailOnLoss = true }, out OfficeConversionReport report);
        Assert.False(report.HasLoss);
        using var saved = new MemoryStream(document.ToBytes());
        using WordprocessingDocument output = WordprocessingDocument.Open(saved, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(output));

        MainDocumentPart main = output.MainDocumentPart;
        Style bodyStyle = main.StyleDefinitionsPart.Styles.Elements<Style>().Single(style => style.StyleId == "MigrationBody");
        Assert.Equal("Normal", bodyStyle.BasedOn.Val.Value);
        Assert.Equal("276", bodyStyle.StyleParagraphProperties.SpacingBetweenLines.Line.Value);
        Assert.Equal("Calibri", bodyStyle.StyleRunProperties.RunFonts.Ascii.Value);
        Assert.Equal("22", bodyStyle.StyleRunProperties.FontSize.Val.Value);
        Assert.Equal("003366", bodyStyle.StyleRunProperties.Color.Val.Value);

        Paragraph paragraph = main.Document.Body.Elements<Paragraph>().First();
        Assert.Equal("MigrationBody", paragraph.ParagraphProperties.ParagraphStyleId.Val.Value);
        Assert.Equal("360", (paragraph.ParagraphProperties.Indentation.Left ?? paragraph.ParagraphProperties.Indentation.Start).Value);
        Assert.Equal(JustificationValues.Both, paragraph.ParagraphProperties.Justification.Val.Value);
        Run title = paragraph.GetFirstChild<Run>();
        Assert.Equal("Emphasis", title.RunProperties.RunStyle.Val.Value);
        Assert.Equal(UnderlineValues.Double, title.RunProperties.Underline.Val.Value);
        Assert.Equal(VerticalPositionValues.Superscript, title.RunProperties.VerticalTextAlignment.Val.Value);
        Assert.Equal("Lead ", title.InnerText);
        Hyperlink link = Assert.Single(paragraph.Elements<Hyperlink>());
        Assert.Equal("Details", link.InnerText);
        Assert.Equal("https://example.invalid/report#table", main.HyperlinkRelationships.Single().Uri.AbsoluteUri);
        Assert.Equal(main.HyperlinkRelationships.Single().Id, link.Id.Value);

        Table table = Assert.Single(main.Document.Body.Elements<Table>());
        Assert.Equal(new[] { "1800", "2400", "3000" }, table.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(column => column.Width.Value));
        TableCell firstCell = table.Elements<TableRow>().First().Elements<TableCell>().First();
        Assert.Equal(2, firstCell.TableCellProperties.GridSpan.Val.Value);
        Assert.Equal("4200", firstCell.TableCellProperties.TableCellWidth.Width.Value);
        Assert.Equal("DDEEFF", firstCell.TableCellProperties.Shading.Fill.Value);
        Assert.Equal("Cell one", firstCell.InnerText);

        SectionProperties section = Assert.Single(main.Document.Body.Elements<SectionProperties>());
        Assert.Equal(PageOrientationValues.Landscape, section.GetFirstChild<PageSize>().Orient.Value);
        Assert.Equal("720", section.GetFirstChild<PageMargin>().Left.Value.ToString());
        Assert.Equal(2, section.GetFirstChild<Columns>().ColumnCount.Value);
        Assert.Equal("144", section.GetFirstChild<Columns>().Space.Value);
        HeaderPart header = Assert.Single(main.HeaderParts);
        FooterPart footer = Assert.Single(main.FooterParts);
        Assert.Equal("Header label", header.Header.InnerText);
        Assert.Equal("Footer label", footer.Footer.InnerText);
        Assert.Equal(header.Uri, main.GetPartById(section.GetFirstChild<HeaderReference>().Id).Uri);
        Assert.Equal(footer.Uri, main.GetPartById(section.GetFirstChild<FooterReference>().Id).Uri);
        Assert.Single(main.ImageParts);
        Assert.Single(header.ImageParts);
        Assert.Equal(PixelBytes, ReadBytes(main.ImageParts.Single()));
        Assert.Equal(PixelBytes, ReadBytes(header.ImageParts.Single()));
        Assert.Contains("width:12pt;height:6pt", main.Document.Descendants<DocumentFormat.OpenXml.Vml.Shape>().Single().Style.Value);
    }

    [Theory]
    [InlineData("<w:spacing w:line=\"240\" w:unsupported=\"1\"/>", "}unsupported")]
    [InlineData("<w:rPr><w:color w:val=\"003366\" w:themeColor=\"accent1\"/></w:rPr>", "}themeColor")]
    [InlineData("<w:rPr><w:spacing w:before=\"120\"/></w:rPr>", "}before")]
    public void UnsupportedFormattingAttributeRemainsAReportedStrictLoss(string properties, string location) {
        string xml = "<w:wordDocument xmlns:w=\"http://schemas.microsoft.com/office/word/2003/wordml\"><w:body>" +
            "<w:p><w:pPr>" + properties + "</w:pPr><w:r><w:t>Visible</w:t></w:r></w:p></w:body></w:wordDocument>";
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        using WordDocument document = WordDocument.Load(input, null, out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, finding => finding.Location?.EndsWith(location, StringComparison.Ordinal) == true &&
            finding.LossKind == OfficeConversionLossKind.Omission);
        Assert.Throws<OfficeConversionException>(() => WordDocument.Load(input,
            new WordLoadOptions { FailOnLoss = true, ReportUnsupportedContent = false }));
    }

    [Theory]
    [InlineData("word-rich-native-2003-xml.xml", true)]
    [InlineData("word-rich-native-flat-opc.xml", false)]
    public void NativeWordProducerPreservesRichBodyAndReferencedStories(string name, bool projected) {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "WordXmlMigration", name);
        using WordDocument document = WordDocument.Load(path, null, out OfficeConversionReport report);
        Assert.Equal(projected, report.HasLoss);
        using var saved = new MemoryStream(document.ToBytes());
        using WordprocessingDocument output = WordprocessingDocument.Open(saved, false);
        MainDocumentPart main = output.MainDocumentPart;
        byte[] image = ReadBytes(Assert.Single(main.ImageParts));
        using var sha = System.Security.Cryptography.SHA256.Create();
        Assert.Equal("5E3D382DB4DD83D59AA5742793AD6B7903409E865C83BCBC54835049F043BC15", BitConverter.ToString(sha.ComputeHash(image)).Replace("-", ""));
        Assert.Equal(2, projected ? main.Document.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Count() : main.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Count());
        Assert.Equal(new[] { "2200", "2200", "2200" }, Assert.Single(main.Document.Descendants<TableGrid>()).Elements<GridColumn>().Select(column => column.Width.Value));
        TableCell merged = main.Document.Descendants<Table>().Single().Elements<TableRow>().First().Elements<TableCell>().First();
        Assert.Equal(2, merged.TableCellProperties.GridSpan.Val.Value);
        Assert.Equal("4400", merged.TableCellProperties.TableCellWidth.Width.Value);
        Assert.Equal("CCEEDD", merged.TableCellProperties.Shading.Fill.Value);
        Assert.Equal(8U, merged.TableCellProperties.TableCellBorders.TopBorder.Size.Value);

        Style body = main.StyleDefinitionsPart.Styles.Elements<Style>().Single(style => style.StyleId == "MigrationBody");
        Assert.Equal("Normal", body.BasedOn.Val.Value);
        Assert.Equal("Arial", body.StyleRunProperties.RunFonts.Ascii.Value);
        Assert.Equal("22", body.StyleRunProperties.FontSize.Val.Value);
        Style accent = main.StyleDefinitionsPart.Styles.Elements<Style>().Single(style => style.StyleId == "MigrationAccent");
        Assert.Equal("663399", accent.StyleRunProperties.Color.Val.Value);
        Assert.Equal("30", accent.StyleRunProperties.Spacing.Val.Value.ToString());
        Assert.NotNull(accent.StyleRunProperties.Bold);
        Assert.NotNull(accent.StyleRunProperties.Italic);
        Assert.Equal("https://example.invalid/migration-reference", main.HyperlinkRelationships.Single().Uri.AbsoluteUri);

        SectionProperties section = Assert.Single(main.Document.Descendants<SectionProperties>());
        Assert.Equal(6, section.Elements().Count(element => element is HeaderReference || element is FooterReference));
        HeaderReference headerRef = section.Elements<HeaderReference>().Single(reference => reference.Type == HeaderFooterValues.Default);
        HeaderPart header = Assert.IsType<HeaderPart>(main.GetPartById(headerRef.Id));
        Assert.Contains("Header migration marker", header.Header.InnerText);
        Assert.Single(header.Header.Descendants<Table>());
        Assert.Equal(image, ReadBytes(Assert.Single(header.ImageParts)));
        FooterReference footerRef = section.Elements<FooterReference>().Single(reference => reference.Type == HeaderFooterValues.Default);
        FooterPart footer = Assert.IsType<FooterPart>(main.GetPartById(footerRef.Id));
        Assert.Equal(" PAGE  \\* MERGEFORMAT ", Assert.Single(footer.Footer.Descendants<FieldCode>()).InnerText);
        Assert.Equal(new[] { FieldCharValues.Begin, FieldCharValues.Separate, FieldCharValues.End },
            footer.Footer.Descendants<FieldChar>().Select(field => field.FieldCharType.Value));
        Assert.Equal(16839U, section.GetFirstChild<PageSize>().Width.Value);
        Assert.Equal(11907U, section.GetFirstChild<PageSize>().Height.Value);
        PageMargin margin = section.GetFirstChild<PageMargin>();
        Assert.Equal(new[] { "840", "960", "720", "1080", "360", "400" },
            new[] { margin.Top.InnerText, margin.Bottom.InnerText, margin.Left.InnerText, margin.Right.InnerText, margin.Header.InnerText, margin.Footer.InnerText });
        Assert.Equal(2, section.GetFirstChild<Columns>().ColumnCount.Value);
        Assert.Equal("480", section.GetFirstChild<Columns>().Space.Value);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(output));
    }

    private static byte[] ReadBytes(ImagePart part) {
        using Stream input = part.GetStream();
        using var bytes = new MemoryStream();
        input.CopyTo(bytes);
        return bytes.ToArray();
    }

    private static readonly byte[] PixelBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=");

    private static string RichDocument() => """
        <w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml"
          xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint">
          <w:fonts><w:defaultFonts w:ascii="Calibri" w:h-ansi="Calibri" w:fareast="MS Mincho" w:cs="Arial"/></w:fonts>
          <w:styles>
            <w:style w:type="paragraph" w:default="on" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
            <w:style w:type="paragraph" w:styleId="MigrationBody"><w:name w:val="Migration body"/><w:basedOn w:val="Normal"/>
              <w:pPr><w:spacing w:before="120" w:after="80" w:line="276" w:line-rule="auto"/></w:pPr>
              <w:rPr><w:rFonts w:ascii="Calibri" w:h-ansi="Calibri"/><w:color w:val="003366"/><w:sz w:val="22"/></w:rPr>
            </w:style>
            <w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:rPr><w:i/></w:rPr></w:style>
          </w:styles>
          <w:body><wx:sect>
            <w:p><w:pPr><w:pStyle w:val="MigrationBody"/><w:ind w:left="360" w:first-line="180"/><w:jc w:val="both"/></w:pPr>
              <w:r><w:rPr><w:rStyle w:val="Emphasis"/><w:u w:val="double"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>Lead </w:t></w:r>
              <w:hlink w:dest="https://example.invalid/report#table" w:screenTip="Details"><w:r><w:t>Details</w:t></w:r></w:hlink>
            </w:p>
            <w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/><w:tblLayout w:type="Fixed"/></w:tblPr>
              <w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="2400"/><w:gridCol w:w="3000"/></w:tblGrid>
              <w:tr><w:trPr><w:tblHeader/><w:trHeight w:val="360" w:h-rule="atLeast"/></w:trPr>
                <w:tc><w:tcPr><w:tcW w:w="4200" w:type="dxa"/><w:gridSpan w:val="2"/><w:shd w:val="clear" w:fill="DDEEFF"/></w:tcPr><w:p><w:r><w:t>Cell one</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell two</w:t></w:r></w:p></w:tc>
              </w:tr>
            </w:tbl>
            <w:p><w:r><w:pict>
              <w:binData w:name="wordml://pixel.png">iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=</w:binData>
              <v:shape id="Picture1" alt="Body pixel" style="width:12pt;height:6pt"><v:imagedata src="wordml://pixel.png"/></v:shape>
            </w:pict></w:r></w:p>
            <w:sectPr>
              <w:hdr w:type="odd"><w:p><w:r><w:t>Header label</w:t><w:pict><v:shape id="HeaderPicture" style="width:12pt;height:6pt"><v:imagedata src="wordml://pixel.png"/></v:shape></w:pict></w:r></w:p></w:hdr>
              <w:ftr w:type="odd"><w:p><w:r><w:t>Footer label</w:t></w:r></w:p></w:ftr>
              <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="720" w:left="720" w:bottom="720" w:right="720" w:header="360" w:footer="360" w:gutter="0"/>
              <w:cols w:num="2" w:space="144"/><w:docGrid w:line-pitch="360"/>
            </w:sectPr>
          </wx:sect></w:body>
        </w:wordDocument>
        """;
}

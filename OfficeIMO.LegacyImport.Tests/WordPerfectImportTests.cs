using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Legacy;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using static OfficeIMO.LegacyImport.Tests.WordPerfectFixture;

namespace OfficeIMO.LegacyImport.Tests;

public sealed class WordPerfectImportTests {
    private static LegacyWordImportResult Import(byte[] source, OfficeLegacyImportLimits? limits = null) =>
        LegacyWordImporter.Import(source, new LegacyWordImportOptions { RequireStructured = true, Limits = limits ?? new OfficeLegacyImportLimits() });

    [Theory]
    [InlineData("WP5.wp", "wordperfect-5-records", "Page 1", "page 2")]
    [InlineData("WP6.wpd", "wordperfect-6-records", "Foo", "foo")]
    [InlineData("testWordPerfect_5_0.wp", "wordperfect-5-records", "President", "Lucid")]
    [InlineData("testWordPerfect_5_1.wp", "wordperfect-5-records", "REPORT TITLE", "PETROLEUM")]
    [InlineData("testWordPerfect.wpd", "wordperfect-6-records", "APPENDIX", "AND FURTHER")]
    public void ReadsIndependentProducerDocuments(string file, string profile, string first, string later) {
        using LegacyWordImportResult result = Import(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "WordPerfect", file)));
        Assert.Equal(profile, result.Detection.ProfileId);
        Assert.Equal(OfficeLegacyImportQuality.Structured, result.Report.Quality);
        Assert.Contains(first, result.PlainText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(later, result.PlainText, StringComparison.OrdinalIgnoreCase);
        using var docx = new MemoryStream(); result.Value.Save(docx); docx.Position = 0;
        using WordDocument loaded = WordDocument.Load(docx);
        Assert.Contains(first, string.Concat(loaded.Paragraphs.Select(paragraph => paragraph.Text)), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RetainsIndependentHeadersAndPageGeometryAcrossSections() {
        using LegacyWordImportResult result = Import(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "WordPerfect", "WP6.wpd")));
        Assert.Equal(2, result.Content.Sections.Count);
        Assert.InRange(result.Content.Sections[0].PageWidthPoints!.Value, 595, 596);
        Assert.InRange(result.Content.Sections[0].MarginLeftPoints!.Value, 70, 72);
        Assert.Equal("Header Type A", string.Concat(result.Value.Sections[0].Header.Default!.Paragraphs.Select(paragraph => paragraph.Text)));
        Assert.Equal("Header Type A but changed", string.Concat(result.Value.Sections[1].Header.Default!.Paragraphs.Select(paragraph => paragraph.Text)));
        Assert.Equal(WordSectionBreakType.NextPage, result.Value.Sections[1].BreakType);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(result.Value.OpenXmlDocument));
    }

    [Theory]
    [InlineData(5)]
    [InlineData(6)]
    public void ProjectsFormattingAndAnchoredNotesInBodyOrder(int version) {
        byte[] bold = version == 6 ? new byte[] { 0xf2, 12, 0xf2 } : new byte[] { 0xc3, 12, 0xc3 };
        byte[] off = version == 6 ? new byte[] { 0xf3, 12, 0xf3 } : new byte[] { 0xc4, 12, 0xc4 };
        byte[] source = version == 6 ? Document6(Join(Text("Before"), bold, Text("bold"), off,
                Function6(0xd7, 0, null, 1), Text("99"), Function6(0xd7, 1), Text("After")),
            (8, TextPacket(Join(bold, Text("Note"), off)))) :
            Document5(Join(Text("Before"), bold, Text("bold"), off,
                Function5(0xd6, 0, Join(new byte[15], Text("Note"))), Text("After")));
        using LegacyWordImportResult result = Import(source);
        Assert.Equal("BeforeboldAfter", result.Content.Paragraphs[0].Text);
        Assert.Contains(result.Content.Paragraphs[0].Runs, run => run.Bold && run.Text == "bold");
        Assert.Equal(0, result.Content.Paragraphs[0].Runs.Single(run => run.NoteIndex.HasValue).NoteIndex);
        Assert.True(result.Content.Notes[0].IsAnchored);
        Assert.Equal("Note", result.Content.Notes[0].Text);
        var xml = result.Value.OpenXmlDocument.MainDocumentPart!.Document;
        Assert.Single(xml!.Descendants<FootnoteReference>());
        Assert.Contains("Note", result.Value.OpenXmlDocument.MainDocumentPart!.FootnotesPart!.Footnotes!.InnerText);
        Assert.DoesNotContain("Recovered Footnote", xml.InnerText);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(result.Value.OpenXmlDocument));
    }

    [Theory]
    [InlineData(5)]
    [InlineData(6)]
    public void ProjectsTableGridBetweenBodyParagraphs(int version) {
        byte[] source;
        if (version == 6) {
            byte[] column = new byte[17]; Array.Copy(Number(2400), 0, column, 1, 2);
            source = Document6(Join(Text("Before"), new byte[] { 0x87 }, Function6(0xd4, 0x2a, new byte[6]),
                Function6(0xd4, 0x2c, column), Function6(0xd4, 0x2c, column), Function6(0xd4, 0x2b),
                new byte[] { 0xc0 }, Text("A1"), new byte[] { 0xc6 }, Text("B1"), new byte[] { 0xc0 },
                Text("A2"), new byte[] { 0xc6 }, Text("B2"), new byte[] { 0xbd }, Text("After")));
        } else {
            var definition = new byte[58]; Array.Copy(Number(2), 0, definition, 26, 2);
            Array.Copy(Number(2400), 0, definition, 48, 2); Array.Copy(Number(2400), 0, definition, 50, 2);
            var cell = new byte[11]; cell[2] = 1; cell[3] = 1;
            source = Document5(Join(Text("Before"), new byte[] { 0x0a }, Function5(0xd2, 0x0b, definition),
                Function5(0xdc, 1), Function5(0xdc, 0, cell), Text("A1"), Function5(0xdc, 0, cell), Text("B1"),
                Function5(0xdc, 1), Function5(0xdc, 0, cell), Text("A2"), Function5(0xdc, 0, cell), Text("B2"), Function5(0xdc, 2), Text("After")));
        }
        using LegacyWordImportResult result = Import(source);
        Assert.Collection(result.Content.Sections[0].Blocks, block => Assert.IsType<LegacyWordParagraphContent>(block),
            block => Assert.Equal(2, Assert.IsType<LegacyWordTableContent>(block).Rows.Count), block => Assert.IsType<LegacyWordParagraphContent>(block));
        WordTable table = Assert.Single(result.Value.Tables);
        Assert.Equal("A1", string.Concat(table.Rows[0].Cells[0].Paragraphs.Select(paragraph => paragraph.Text)));
        Assert.Equal("B2", string.Concat(table.Rows[1].Cells[1].Paragraphs.Select(paragraph => paragraph.Text)));
        Assert.Equal(2880, table.Rows[0].Cells[0].Width);
        Assert.Equal("Before\nA1\nB1\nA2\nB2\nAfter", result.PlainText);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(result.Value.OpenXmlDocument));
    }

    [Fact]
    public void ExcludesDeletedTextAndRejectsUnbalancedUndo() {
        byte[] begin = { 0xf1, 0, 1, 0, 0xf1 }, end = { 0xf1, 1, 2, 0, 0xf1 };
        using LegacyWordImportResult result = Import(Document6(Join(Text("Visible"), begin, Text("Deleted"), end, Text("Tail"))));
        Assert.Equal("VisibleTail", result.PlainText);
        Assert.Throws<InvalidDataException>(() => Import(Document6(Join(Text("Visible"), begin, Text("Deleted")))));
        Assert.Throws<InvalidDataException>(() => Import(Document6(end)));
        using LegacyWordImportResult nested = Import(Document6(Join(begin, Text("Deleted"), new byte[] { 0xf1, 2, 3, 0, 0xf1 },
            Text("Restored"), new byte[] { 0xf1, 3, 4, 0, 0xf1 }, Text("Deleted"), end)));
        Assert.Equal("Restored", nested.PlainText);
    }

    [Fact]
    public void RejectsBrokenEnvelopesRecursiveStoriesAndEncryptedDocuments() {
        byte[] function = Function6(0xd3, 5, new byte[] { 2 }); function[^1] = 0xd4;
        Assert.Throws<InvalidDataException>(() => Import(Document6(function)));
        Assert.Throws<InvalidDataException>(() => Import(Document6(Function6(0xd6, 0, new byte[] { 3 }, 1),
            (8, TextPacket(Function6(0xd6, 0, new byte[] { 3 }, 1))))));
        byte[] encrypted = Document5(Text("Visible")); encrypted[12] = 1;
        Assert.Throws<InvalidDataException>(() => Import(encrypted));
    }

    [Fact]
    public void SharesLimitsAcrossBodyAndReferencedStories() {
        byte[] source = Document6(Join(Text("Body"), Function6(0xd7, 0, null, 1), Function6(0xd7, 1)), (8, TextPacket(Text("Note"))));
        Assert.Throws<InvalidDataException>(() => Import(source, new OfficeLegacyImportLimits { MaxTextCharacters = 7 }));
        Assert.Throws<InvalidDataException>(() => Import(source, new OfficeLegacyImportLimits { MaxRecords = 2 }));
        Assert.Throws<InvalidDataException>(() => Import(source, new OfficeLegacyImportLimits { MaxItems = 2 }));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => LegacyWordImporter.Import(source, cancellationToken: cancellation.Token));
    }

    [Fact]
    public void RecoversIndependentWpgVectorsAsAnInlinePngWithImmutableSourceData() {
        byte[] wpg = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "WordPerfect", "WPG1.wpg"));
        byte[] source = GraphicDocument(wpg);
        using LegacyWordImportResult result = Import(source);
        LegacyWordImageContent image = result.Content.Paragraphs[0].Runs.Single(run => run.Image != null).Image!;
        Assert.Equal(wpg, image.GetSourceBytes());
        byte[] copy = image.GetSourceBytes(); copy[0] = 0;
        Assert.Equal(0xff, image.GetSourceBytes()[0]);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, image.GetPngBytes().Take(8));
        Assert.InRange(image.WidthPoints, 376, 377);
        Assert.InRange(image.HeightPoints, 474, 475);
        Assert.Single(result.Value.Images);
        Assert.Equal("BeforeAfter", result.PlainText);
        Assert.Contains(result.Report.Findings, finding => finding.Code == "WORDPERFECT_GRAPHIC_LAYOUT");
        using var docx = new MemoryStream(); result.Value.Save(docx); docx.Position = 0;
        using WordDocument loaded = WordDocument.Load(docx); Assert.Single(loaded.Images);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(result.Value.OpenXmlDocument));
    }

    [Fact]
    public void RejectsImageExpansionAndResourceBudgetsBeforeProjection() {
        byte[] wpg = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "WordPerfect", "WPG1.wpg"));
        Assert.Throws<InvalidDataException>(() => Import(GraphicDocument(wpg), new OfficeLegacyImportLimits { MaxResourceBytes = wpg.Length - 1 }));
        Assert.Throws<InvalidDataException>(() => Import(GraphicDocument(wpg), new OfficeLegacyImportLimits { MaxImagePixels = 100 }));
        byte[] broken = (byte[])wpg.Clone(); broken[17] = 255;
        Assert.Throws<InvalidDataException>(() => Import(GraphicDocument(broken)));
        byte[] unsupported = (byte[])wpg.Clone(); unsupported[10] = 2;
        using LegacyWordImportResult result = Import(GraphicDocument(unsupported));
        Assert.Empty(result.Value.Images);
        Assert.Contains(result.Report.Findings, finding => finding.Code == "WORDPERFECT_GRAPHIC_PROFILE");
        Assert.Equal(OfficeLegacyInertContentKind.None, result.Report.InertContent);
    }

    [Fact]
    public void PdfKeepsAnInlineGraphicBetweenItsSurroundingText() {
        using LegacyWordImportResult result = Import(GraphicDocument(File.ReadAllBytes(
            Path.Combine(AppContext.BaseDirectory, "Fixtures", "WordPerfect", "WPG1.wpg"))));
        var page = Assert.Single(PdfDocument.Load(result.Value.ToPdfBytes()).Read().Pages);
        var before = Assert.Single(page.TextBlocks.SelectMany(block => block.Spans), span => span.Text == "Before");
        var after = Assert.Single(page.TextBlocks.SelectMany(block => block.Spans), span => span.Text == "After");
        var image = Assert.Single(page.Images).PrimaryPlacement!;
        Assert.True(before.X < image.X);
        Assert.True(after.X >= image.X + image.Width - 0.1);
        Assert.Equal(before.Y, after.Y, 2);
    }

    private static byte[] GraphicDocument(byte[] wpg) {
        byte[] source = Document6(Join(Text("Before"), Function6(0xdf, 0, new byte[20], 1, 2), Text("After")),
            (0x41, new byte[32]), (0x40, Join(Number(1), Number(3), Number(1), Text("fixture.wpg\0"))), (0x6f, wpg));
        source[512 + 2 * 14] = 1; // Graphics-filename packet contains a child-ID directory.
        return source;
    }
}

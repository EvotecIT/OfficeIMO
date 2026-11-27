using System.IO.Compression;
using Microsoft.Extensions.Documents;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader.All;
using OfficeIMO.Word;
using Xunit;
using static OfficeIMO.Reader.DocumentExtraction.Tests.ExtractionMappingTests;

namespace OfficeIMO.Reader.DocumentExtraction.Tests;

/// <summary>Real documents produced by the OfficeIMO engines (or written by hand for HTML and EPUB) through the client.</summary>
public sealed class ExtractionFixtureTests {
    private static readonly OfficeIMODocumentExtractionClient Client = new(builder => builder.AddAllOfficeIMOHandlers());

    [Fact]
    public async Task WordDocumentKeepsHeadingsListsTablesAndImages() {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create()) {
            document.AddParagraph("Policy").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Applies to all staff.");
            WordList list = document.AddList(WordListStyle.Bulleted);
            list.AddItem("First rule");
            list.AddItem("Second rule");
            WordTable totals = document.AddTable(2, 2);
            totals.Rows[0].Cells[0].Paragraphs[0].Text = "Region";
            totals.Rows[0].Cells[1].Paragraphs[0].Text = "Total";
            totals.Rows[1].Cells[0].Paragraphs[0].Text = "North";
            totals.Rows[1].Cells[1].Paragraphs[0].Text = "10";
            WordImage image = document.AddParagraph("Logo:").AddImage(ImagePath);
            image.AltText = "Company logo";
            bytes = document.ToBytes();
        }

        ExtractionDocument extracted = await ExtractAsync(bytes, "policy.docx");

        ExtractionDocumentSection policy = extracted.Sections[0];
        var heading = Assert.IsType<ExtractionDocumentTextElement>(policy.Elements[0]);
        Assert.Equal(("heading", 1), (heading.Role, heading.HeadingLevel));
        ExtractionDocumentElement[] content = extracted.EnumerateContent().ToArray();
        Assert.Contains(content, element => element is ExtractionDocumentTextElement { Role: null } && Text(element) == "Applies to all staff.");
        Assert.Equal(2, content.OfType<ExtractionDocumentTextElement>().Count(element => element.Role == "listItem"));
        // The table has no header row in Word, so no cell is a column header, and its text projection is not repeated.
        ExtractionDocumentTableElement table = Assert.Single(content.OfType<ExtractionDocumentTableElement>());
        Assert.Equal((2, 2), (table.Cells.RowCount, table.Cells.ColumnCount));
        Assert.Equal("Region", Text(table.Cells[0, 0]!.Content!));
        Assert.All(table.Cells, cell => Assert.Null(cell.Role));
        Assert.DoesNotContain(content, element => element is ExtractionDocumentTextElement && Text(element).Contains("Region"));
        // Word reports no position for images, so the image sits in its own trailing section.
        ExtractionDocumentImageElement logo = Assert.IsType<ExtractionDocumentImageElement>(Assert.Single(extracted.Sections[^1].Elements));
        Assert.Equal("Company logo", logo.AlternativeText);
        Assert.IsType<OfficeDocumentReadResult>(extracted.RawRepresentation);
    }

    [Fact]
    public async Task SpreadsheetBecomesATableOnItsSheetFrame() {
        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create()) {
            ExcelSheet sheet = document.AddSheet("Budget");
            sheet.Cell(1, 1).SetValue("Item");
            sheet.Cell(1, 2).SetValue("Cost");
            sheet.Cell(2, 1).SetValue("Paper");
            sheet.Cell(2, 2).SetValue(12);
            bytes = document.ToBytes();
        }

        ExtractionDocument extracted = await ExtractAsync(bytes, "budget.xlsx");

        ExtractionDocumentTableElement table = Assert.Single(extracted.EnumerateContent().OfType<ExtractionDocumentTableElement>());
        Assert.Contains(table.Cells, cell => cell.Content is ExtractionDocumentTextElement text && Text(text) == "Paper");
        ExtractionDocumentSourcePage frame = Assert.Single(table.Sources).Page;
        Assert.Equal("sheet-Budget", frame.Identifier);
        Assert.Equal("Budget", table.Content!.AdditionalProperties![OfficeIMOExtractionPropertyNames.Sheet]);
    }

    [Fact]
    public async Task PresentationSlidesBecomeNumberedFrames() {
        byte[] bytes;
        using (PowerPointDocument presentation = PowerPointDocument.Create()) {
            presentation.AddSlide().AddTextBox("Intro slide");
            presentation.AddSlide().AddTextBox("Details slide");
            bytes = presentation.ToBytes();
        }

        ExtractionDocument extracted = await ExtractAsync(bytes, "deck.pptx");

        ExtractionDocumentElement details = extracted.EnumerateContent().Single(element => element is ExtractionDocumentTextElement && Text(element).Contains("Details slide"));
        ExtractionDocumentSourcePage frame = Assert.Single(details.Sources).Page;
        Assert.Equal(2, frame.PageNumber);
        Assert.Contains(extracted.SourcePages, page => page.PageNumber == 1);
    }

    [Fact]
    public async Task PdfHeadingsOpenSectionsOnSizedPages() {
        PdfDocument pdf = PdfDocument.Create(
            builder => builder.Content(content => content.H1("PDF Title").AddParagraph("This is a PDF body.")),
            new PdfCreateOptions { PageSetup = { Size = new PdfPageSize(OfficeLength.Points(420), OfficeLength.Points(360)) } });

        ExtractionDocument extracted = await ExtractAsync(pdf.ToBytes(), "report.pdf");

        ExtractionDocumentSection section = extracted.Sections.Single(candidate => candidate.Elements.FirstOrDefault() is ExtractionDocumentTextElement { Role: "heading" });
        Assert.Equal("PDF Title", Text(section.Elements[0]));
        Assert.Contains(extracted.EnumerateContent(), element => element is ExtractionDocumentTextElement && Text(element).Contains("This is a PDF body."));
        ExtractionDocumentSourcePage page = Assert.Single(extracted.SourcePages);
        Assert.Equal(1, page.PageNumber);
        Assert.Equal(new ExtractionDocumentPageDimensions(420, 360), page.Dimensions);
        Assert.Equal(ExtractionDocumentCoordinateOrigin.BottomLeft, page.CoordinateOrigin);
        Assert.Same(page, Assert.Single(section.Elements[0].Sources).Page);
    }

    [Fact]
    public async Task HtmlKeepsHeadingHierarchyListsAndTableHeaders() {
        const string html = "<html><head><title>Guide</title></head><body>" +
            "<h1>Guide</h1><p>Intro text.</p><h2>Steps</h2><ul><li>Open</li><li>Close</li></ul>" +
            "<table><tr><th>Name</th><th>Qty</th></tr><tr><td>Bolt</td><td>4</td></tr></table>" +
            "<img src=\"data:image/gif;base64,R0lGODlhAQABAJAAAAAAAP///ywAAAAAAQABAAACAkwBADs=\" alt=\"Dot\"/></body></html>";

        ExtractionDocument extracted = await ExtractAsync(System.Text.Encoding.UTF8.GetBytes(html), "guide.html");

        ExtractionDocumentSection guide = extracted.Sections.Single(section => section.Elements.FirstOrDefault() is ExtractionDocumentTextElement { Role: "heading" } heading && Text(heading) == "Guide");
        ExtractionDocumentSection steps = Assert.Single(guide.Elements.OfType<ExtractionDocumentSection>());
        Assert.Equal(2, Assert.IsType<ExtractionDocumentTextElement>(steps.Elements[0]).HeadingLevel);
        Assert.Equal(2, steps.Elements.OfType<ExtractionDocumentTextElement>().Count(element => element.Role == "listItem"));
        // The reader also projects the table from its text chunk; only one table may come out.
        ExtractionDocumentTableElement table = Assert.Single(extracted.EnumerateContent().OfType<ExtractionDocumentTableElement>());
        Assert.Contains(table, steps.Elements);
        Assert.Equal("columnHeader", table.Cells[0, 0]!.Role);
        Assert.Equal("Name", Text(table.Cells[0, 0]!.Content!));
        Assert.Equal("Bolt", Text(table.Cells[1, 0]!.Content!));
        ExtractionDocumentImageElement image = Assert.Single(steps.Elements.OfType<ExtractionDocumentImageElement>());
        Assert.Equal("Dot", image.AlternativeText);
        Assert.DoesNotContain(extracted.EnumerateContent(), element => element is ExtractionDocumentTextElement { Role: "image" });
    }

    [Fact]
    public async Task EpubChaptersBecomeSections() {
        ExtractionDocument extracted = await ExtractAsync(BuildEpub(), "book.epub");

        ExtractionDocumentTextElement[] headings = extracted.EnumerateContent().OfType<ExtractionDocumentTextElement>().Where(element => element.Role == "heading").ToArray();
        Assert.Equal(new[] { "One", "Two" }, headings.Select(Text).ToArray());
        Assert.Contains(extracted.EnumerateContent(), element => element is ExtractionDocumentTextElement && Text(element) == "Second chapter text.");
        Assert.Equal("Demo Book", extracted.Metadata["title"]);
    }

    [Fact]
    public async Task MediaTypeSelectsTheHandlerWhenTheIdentifierHasNoExtension() {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create()) {
            document.AddParagraph("Uploaded body");
            bytes = document.ToBytes();
        }

        ExtractionDocument extracted = await ExtractAsync(bytes, "upload-42", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");

        Assert.Equal("upload-42", extracted.Identifier);
        Assert.Equal("Word", extracted.Metadata[OfficeIMOExtractionPropertyNames.InputKind]);
        Assert.Contains(extracted.EnumerateContent(), element => Text(element) == "Uploaded body");
    }

    internal static async Task<ExtractionDocument> ExtractAsync(byte[] bytes, string identifier, string? mediaType = null) {
        using var stream = new MemoryStream(bytes, writable: false);
        ExtractionDocument document = await Client.ExtractAsync(stream, identifier, mediaType ?? string.Empty);
        Assert.True(stream.CanRead, "The client must not dispose the caller's stream.");
        return document;
    }

    private static string ImagePath => Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");

    private static byte[] BuildEpub() {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true)) {
            Write(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
            Write(archive, "META-INF/container.xml",
                "<?xml version=\"1.0\"?><container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">" +
                "<rootfiles><rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles></container>");
            Write(archive, "OEBPS/content.opf",
                "<?xml version=\"1.0\"?><package version=\"3.0\" unique-identifier=\"id\" xmlns=\"http://www.idpf.org/2007/opf\">" +
                "<metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>Demo Book</dc:title><dc:language>en</dc:language><dc:identifier id=\"id\">book-1</dc:identifier></metadata>" +
                "<manifest><item id=\"nav\" href=\"nav.xhtml\" media-type=\"application/xhtml+xml\" properties=\"nav\"/>" +
                "<item id=\"ch1\" href=\"ch1.xhtml\" media-type=\"application/xhtml+xml\"/><item id=\"ch2\" href=\"ch2.xhtml\" media-type=\"application/xhtml+xml\"/></manifest>" +
                "<spine><itemref idref=\"ch1\"/><itemref idref=\"ch2\"/></spine></package>");
            Write(archive, "OEBPS/nav.xhtml",
                "<?xml version=\"1.0\"?><html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:epub=\"http://www.idpf.org/2007/ops\"><body><nav epub:type=\"toc\"><ol>" +
                "<li><a href=\"ch1.xhtml\">One</a></li><li><a href=\"ch2.xhtml\">Two</a></li></ol></nav></body></html>");
            Write(archive, "OEBPS/ch1.xhtml",
                "<?xml version=\"1.0\"?><html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>One</title></head><body><h1>One</h1><p>First chapter text.</p></body></html>");
            Write(archive, "OEBPS/ch2.xhtml",
                "<?xml version=\"1.0\"?><html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Two</title></head><body><h1>Two</h1><p>Second chapter text.</p></body></html>");
        }
        return buffer.ToArray();
    }

    private static void Write(ZipArchive archive, string path, string content, CompressionLevel level = CompressionLevel.Optimal) {
        using Stream entry = archive.CreateEntry(path, level).Open();
        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(content);
        entry.Write(bytes, 0, bytes.Length);
    }
}

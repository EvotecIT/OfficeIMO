using OfficeIMO.Email;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Reader.Benchmarks.Comparison;

internal static class ReaderComparisonCorpus {
    private static readonly DateTime FixedPackageTimestamp =
        new DateTime(2026, 7, 15, 12, 0, 0, DateTimeKind.Utc);

    private const string TinyPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

    public static IReadOnlyList<ReaderComparisonCase> Create() => new[] {
        CreateWord(),
        CreateExcel(),
        CreatePowerPoint(),
        CreatePdf(),
        CreateHtml(),
        CreateCsv(),
        CreateMsg(),
        CreateEpub(),
        CreateZip(),
        CreateMalformedPdf()
    };

    private static ReaderComparisonCase CreateWord() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            document.AddParagraph("Evidence policy").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Policy portal: ")
                .AddHyperLink("Open policy portal", new Uri("https://example.com/policy"), addStyle: true);
            document.AddParagraph("Footnote anchor")
                .AddFootNote("Footnote retention marker");
            WordTable table = document.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Status";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Retention marker";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "Ready";
            document.BuiltinDocumentProperties.Created = FixedPackageTimestamp;
            document.BuiltinDocumentProperties.Modified = FixedPackageTimestamp;
            document.Save();
        }

        return Case("docx", "evidence-policy.docx", NormalizeOfficePackage(stream.ToArray()),
            Probe("heading", ReaderComparisonProbeKind.MarkdownHeading, "Evidence policy"),
            Probe(
                "link",
                ReaderComparisonProbeKind.MarkdownLink,
                "Open policy portal",
                "https://example.com/policy"),
            Probe("footnote", ReaderComparisonProbeKind.ContainsText, "Footnote retention marker"),
            Probe("table-text", ReaderComparisonProbeKind.MarkdownTable, "Retention marker"),
            Probe("rich-table", ReaderComparisonProbeKind.RichTable),
            Probe("rich-link", ReaderComparisonProbeKind.RichLink),
            Probe("heading-location", ReaderComparisonProbeKind.LocationHeading, "Evidence policy"));
    }

    private static ReaderComparisonCase CreateExcel() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-reader-evidence-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Evidence");
                sheet.Cell(1, 1, "Control");
                sheet.Cell(1, 2, "Owner");
                sheet.Cell(1, 3, "Status");
                sheet.Cell(2, 1, "Spreadsheet retention marker");
                sheet.Cell(2, 2, "Reader team");
                sheet.Cell(2, 3, "Ready");
                document.BuiltinDocumentProperties.Created = FixedPackageTimestamp;
                document.BuiltinDocumentProperties.Modified = FixedPackageTimestamp;
                document.Save();
            }

            return Case("xlsx", "evidence-workbook.xlsx", NormalizeOfficePackage(File.ReadAllBytes(path)),
                Probe("table-text", ReaderComparisonProbeKind.ContainsText, "Spreadsheet retention marker"),
                Probe("markdown-table", ReaderComparisonProbeKind.MarkdownTable, "Spreadsheet retention marker"),
                Probe("rich-table", ReaderComparisonProbeKind.RichTable),
                Probe("sheet-location", ReaderComparisonProbeKind.LocationSheet, "Evidence"));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static ReaderComparisonCase CreatePowerPoint() {
        using var stream = new MemoryStream();
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox text = slide.AddTextBox("Quarterly evidence");
            text.AddBullet("Presentation list marker");
            PowerPointTable table = slide.AddTable(2, 2);
            table.GetCell(0, 0).Text = "Measure";
            table.GetCell(0, 1).Text = "Value";
            table.GetCell(1, 0).Text = "Presentation table marker";
            table.GetCell(1, 1).Text = "42";
            using var image = new MemoryStream(Convert.FromBase64String(TinyPngBase64));
            slide.AddPicture(image, OfficeIMO.PowerPoint.ImagePartType.Png);
            slide.Notes.Text = "Presenter notes retention marker";
            presentation.Save();
            presentation.BuiltinDocumentProperties.Created = FixedPackageTimestamp;
            presentation.BuiltinDocumentProperties.Modified = FixedPackageTimestamp;
        }

        return Case("pptx", "evidence-deck.pptx", NormalizeOfficePackage(stream.ToArray()),
            Probe("title", ReaderComparisonProbeKind.ContainsText, "Quarterly evidence"),
            Probe("list", ReaderComparisonProbeKind.MarkdownListItem, "Presentation list marker"),
            Probe("table", ReaderComparisonProbeKind.MarkdownTable, "Presentation table marker"),
            Probe("notes", ReaderComparisonProbeKind.ContainsText, "Presenter notes retention marker"),
            Probe("rich-table", ReaderComparisonProbeKind.RichTable),
            Probe("rich-asset", ReaderComparisonProbeKind.RichAsset),
            Probe("slide-location", ReaderComparisonProbeKind.LocationSlide, expectedSlide: 1));
    }

    private static ReaderComparisonCase CreatePdf() {
        PdfDocument document = PdfDocument.Create();
        document.H1("PDF evidence heading");
        document.Paragraph(value => value.Text("PDF first-page retention marker."));
        document.PageBreak();
        document.H1("PDF second page");
        document.Paragraph(value => value.Text("PDF second-page retention marker."));

        return Case("pdf", "evidence-report.pdf", document.ToBytes(),
            Probe("heading", ReaderComparisonProbeKind.MarkdownHeading, "PDF evidence heading"),
            Probe("page-two", ReaderComparisonProbeKind.ContainsText, "PDF second-page retention marker"),
            Probe("page-location", ReaderComparisonProbeKind.LocationPage, expectedPage: 2));
    }

    private static ReaderComparisonCase CreateHtml() {
        string html = "<!doctype html><html><head><title>Evidence article</title></head><body>" +
            "<h1>HTML evidence heading</h1>" +
            "<ul><li>HTML list retention marker</li></ul>" +
            "<table><thead><tr><th>Control</th><th>Status</th></tr></thead>" +
            "<tbody><tr><td>HTML table retention marker</td><td>Ready</td></tr></tbody></table>" +
            "<p><a href=\"https://example.com/html\">HTML link marker</a></p>" +
            "<img alt=\"HTML image marker\" src=\"data:image/png;base64," + TinyPngBase64 + "\">" +
            "</body></html>";

        return Case("html", "evidence-article.html", Encoding.UTF8.GetBytes(html),
            Probe("heading", ReaderComparisonProbeKind.MarkdownHeading, "HTML evidence heading"),
            Probe("list", ReaderComparisonProbeKind.MarkdownListItem, "HTML list retention marker"),
            Probe("table", ReaderComparisonProbeKind.MarkdownTable, "HTML table retention marker"),
            Probe(
                "link",
                ReaderComparisonProbeKind.MarkdownLink,
                "HTML link marker",
                "https://example.com/html"),
            Probe(
                "image",
                ReaderComparisonProbeKind.MarkdownImage,
                "HTML image marker",
                "data:image/png;base64,"),
            Probe("rich-table", ReaderComparisonProbeKind.RichTable),
            Probe("rich-link", ReaderComparisonProbeKind.RichLink));
    }

    private static ReaderComparisonCase CreateCsv() {
        const string csv = "Control,Owner,Status\nCSV retention marker,Reader team,Ready\n";
        return Case("csv", "evidence-records.csv", Encoding.UTF8.GetBytes(csv),
            Probe("table", ReaderComparisonProbeKind.MarkdownTable, "CSV retention marker"),
            Probe("rich-table", ReaderComparisonProbeKind.RichTable),
            Probe("path-location", ReaderComparisonProbeKind.LocationPath, "evidence-records.csv"));
    }

    private static ReaderComparisonCase CreateMsg() {
        var message = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "MSG evidence subject",
            From = new EmailAddress("sender@example.com", "Evidence Sender"),
            Date = new DateTimeOffset(2026, 7, 15, 12, 0, 0, TimeSpan.Zero)
        };
        message.Body.Text = "MSG body retention marker";
        message.Recipients.Add(new EmailRecipient(
            EmailRecipientKind.To,
            new EmailAddress("reader@example.com", "Reader Recipient")));
        message.Attachments.Add(new EmailAttachment {
            FileName = "evidence.txt",
            ContentType = "text/plain",
            Content = Encoding.UTF8.GetBytes("MSG attachment retention marker"),
            Length = "MSG attachment retention marker".Length
        });
        byte[] bytes = new EmailDocumentWriter().ToBytes(message, EmailFileFormat.OutlookMsg);

        return Case("msg", "evidence-message.msg", bytes,
            Probe("subject", ReaderComparisonProbeKind.ContainsText, "MSG evidence subject"),
            Probe("body", ReaderComparisonProbeKind.ContainsText, "MSG body retention marker"),
            Probe("attachment-name", ReaderComparisonProbeKind.ContainsText, "evidence.txt"),
            Probe("attachment-content", ReaderComparisonProbeKind.ContainsText, "MSG attachment retention marker"),
            Probe("rich-asset", ReaderComparisonProbeKind.RichAsset),
            Probe("path-location", ReaderComparisonProbeKind.LocationPath, "evidence-message.msg"));
    }

    private static ReaderComparisonCase CreateEpub() {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteEntry(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
            WriteEntry(archive, "META-INF/container.xml",
                "<?xml version=\"1.0\"?><container xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles>" +
                "<rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/>" +
                "</rootfiles></container>");
            WriteEntry(archive, "OEBPS/content.opf",
                "<?xml version=\"1.0\"?><package version=\"3.0\" unique-identifier=\"book-id\" xmlns=\"http://www.idpf.org/2007/opf\">" +
                "<metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\">" +
                "<dc:identifier id=\"book-id\">urn:uuid:3f454ae1-3bd5-4ea3-b9b4-8e9c43db295e</dc:identifier>" +
                "<dc:title>Evidence book</dc:title><dc:language>en</dc:language></metadata>" +
                "<manifest><item id=\"chapter\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/>" +
                "<item id=\"nav\" href=\"nav.xhtml\" media-type=\"application/xhtml+xml\" properties=\"nav\"/></manifest>" +
                "<spine><itemref idref=\"chapter\"/></spine></package>");
            WriteEntry(archive, "OEBPS/nav.xhtml",
                "<?xml version=\"1.0\"?><html xmlns=\"http://www.w3.org/1999/xhtml\" " +
                "xmlns:epub=\"http://www.idpf.org/2007/ops\"><head><title>Contents</title></head>" +
                "<body><nav epub:type=\"toc\"><ol><li><a href=\"chapter.xhtml\">Evidence</a></li></ol></nav></body></html>");
            WriteEntry(archive, "OEBPS/chapter.xhtml",
                "<?xml version=\"1.0\"?><html xmlns=\"http://www.w3.org/1999/xhtml\"><body>" +
                "<h1>EPUB evidence heading</h1><ul><li>EPUB list retention marker</li></ul>" +
                "<p><a href=\"https://example.com/epub\">EPUB link marker</a></p></body></html>");
        }

        return Case("epub", "evidence-book.epub", stream.ToArray(),
            Probe("heading", ReaderComparisonProbeKind.MarkdownHeading, "EPUB evidence heading"),
            Probe("list", ReaderComparisonProbeKind.MarkdownListItem, "EPUB list retention marker"),
            Probe(
                "link",
                ReaderComparisonProbeKind.MarkdownLink,
                "EPUB link marker",
                "https://example.com/epub"),
            Probe("path-location", ReaderComparisonProbeKind.LocationPath, "OEBPS/chapter.xhtml"));
    }

    private static ReaderComparisonCase CreateZip() {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteEntry(archive, "docs/evidence.md",
                "# ZIP evidence heading\n\n- ZIP list retention marker\n\n" +
                "[ZIP link marker](https://example.com/zip)\n");
            WriteEntry(archive, "docs/data.csv", "Control,Status\nZIP nested table marker,Ready\n");
        }

        return Case("zip", "evidence-archive.zip", stream.ToArray(),
            Probe("heading", ReaderComparisonProbeKind.MarkdownHeading, "ZIP evidence heading"),
            Probe("list", ReaderComparisonProbeKind.MarkdownListItem, "ZIP list retention marker"),
            Probe(
                "link",
                ReaderComparisonProbeKind.MarkdownLink,
                "ZIP link marker",
                "https://example.com/zip"),
            Probe("nested-table", ReaderComparisonProbeKind.ContainsText, "ZIP nested table marker"),
            Probe("nested-path", ReaderComparisonProbeKind.LocationPath, "docs/evidence.md"),
            Probe("nested-table-path", ReaderComparisonProbeKind.LocationPath, "docs/data.csv"));
    }

    private static ReaderComparisonCase CreateMalformedPdf() => Case(
        "malformed-pdf",
        "malformed.pdf",
        Encoding.ASCII.GetBytes("%PDF-1.7\nnot-a-valid-pdf"),
        Probe("rejected", ReaderComparisonProbeKind.RejectsMalformedInput));

    private static ReaderComparisonCase Case(
        string id,
        string sourceName,
        byte[] bytes,
        params ReaderComparisonProbe[] probes) => new ReaderComparisonCase(id, sourceName, bytes, probes);

    private static ReaderComparisonProbe Probe(
        string id,
        ReaderComparisonProbeKind kind,
        string marker = "",
        string expectedTarget = "",
        int? expectedPage = null,
        int? expectedSlide = null) => new ReaderComparisonProbe(
            id,
            kind,
            marker,
            expectedTarget,
            expectedPage,
            expectedSlide);

    private static byte[] NormalizeOfficePackage(byte[] packageBytes) =>
        ReaderComparisonPackageNormalizer.Normalize(
            packageBytes,
            new DateTimeOffset(FixedPackageTimestamp));

    private static void WriteEntry(
        ZipArchive archive,
        string path,
        string content,
        CompressionLevel compressionLevel = CompressionLevel.Optimal) {
        ZipArchiveEntry entry = archive.CreateEntry(path, compressionLevel);
        entry.LastWriteTime = new DateTimeOffset(FixedPackageTimestamp);
        using Stream entryStream = entry.Open();
        using var writer = new StreamWriter(entryStream, new UTF8Encoding(false));
        writer.Write(content);
    }
}

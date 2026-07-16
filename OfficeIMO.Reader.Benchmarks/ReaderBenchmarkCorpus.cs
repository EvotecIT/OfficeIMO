using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using System.Globalization;
using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Reader.Benchmarks;

internal static class ReaderBenchmarkCorpus {
    private static readonly IReadOnlyDictionary<string, ReaderBenchmarkInput> Inputs = BuildInputs();

    public static IEnumerable<string> Names => Inputs.Keys;

    public static ReaderBenchmarkInput Get(string name) => Inputs[name];

    private static IReadOnlyDictionary<string, ReaderBenchmarkInput> BuildInputs() {
        var inputs = new SortedDictionary<string, ReaderBenchmarkInput>(StringComparer.Ordinal) {
            ["BinaryPowerPoint"] = new ReaderBenchmarkInput("BinaryPowerPoint", "deck.ppt", BuildBinaryPowerPoint()),
            ["Csv"] = Text("Csv", "records.csv", BuildCsv()),
            ["Epub"] = new ReaderBenchmarkInput("Epub", "book.epub", BuildEpub()),
            ["Excel"] = new ReaderBenchmarkInput("Excel", "workbook.xlsx", BuildExcel()),
            ["Html"] = Text("Html", "article.html", BuildHtml()),
            ["Json"] = Text("Json", "records.json", BuildJson()),
            ["Markdown"] = Text("Markdown", "handbook.md", BuildMarkdown()),
            ["Pdf"] = new ReaderBenchmarkInput("Pdf", "report.pdf", BuildPdf()),
            ["PowerPoint"] = new ReaderBenchmarkInput("PowerPoint", "deck.pptx", BuildPowerPoint()),
            ["Rtf"] = Text("Rtf", "notes.rtf", BuildRtf()),
            ["Visio"] = new ReaderBenchmarkInput("Visio", "process.vsdx", BuildVisio()),
            ["Word"] = new ReaderBenchmarkInput("Word", "policy.docx", BuildWord()),
            ["Xml"] = Text("Xml", "records.xml", BuildXml()),
            ["Yaml"] = Text("Yaml", "records.yaml", BuildYaml()),
            ["Zip"] = new ReaderBenchmarkInput("Zip", "archive.zip", BuildZip())
        };
        return inputs;
    }

    private static ReaderBenchmarkInput Text(string name, string sourceName, string text) =>
        new ReaderBenchmarkInput(name, sourceName, Encoding.UTF8.GetBytes(text));

    private static string BuildMarkdown() {
        var builder = new StringBuilder();
        builder.AppendLine("# Reader benchmark handbook");
        for (int section = 1; section <= 80; section++) {
            builder.AppendLine();
            builder.Append("## Section ").AppendLine(section.ToString());
            builder.AppendLine();
            builder.Append("This representative section contains document ingestion guidance, stable citations, and bounded processing notes for item ")
                .Append(section).AppendLine(".");
            builder.AppendLine();
            builder.AppendLine("| Name | Value | Status |");
            builder.AppendLine("| --- | ---: | --- |");
            for (int row = 1; row <= 5; row++) {
                builder.Append("| Item ").Append(row).Append(" | ").Append(section * row).AppendLine(" | Ready |");
            }
        }
        return builder.ToString();
    }

    private static string BuildCsv() {
        var builder = new StringBuilder("Id,Name,Amount,Active,RecordedUtc\n");
        for (int row = 1; row <= 2500; row++) {
            builder.Append(row).Append(",Item ").Append(row).Append(',')
                .Append((row * 1.25m).ToString(CultureInfo.InvariantCulture)).Append(',').Append(row % 2 == 0 ? "true" : "false")
                .Append(",2026-07-10T08:00:00Z\n");
        }
        return builder.ToString();
    }

    private static string BuildJson() {
        var builder = new StringBuilder("[\n");
        for (int row = 1; row <= 1200; row++) {
            if (row > 1) builder.AppendLine(",");
            builder.Append("  {\"id\":").Append(row)
                .Append(",\"name\":\"Item ").Append(row)
                .Append("\",\"active\":").Append(row % 2 == 0 ? "true" : "false")
                .Append(",\"tags\":[\"reader\",\"benchmark\"]}");
        }
        return builder.AppendLine().Append(']').ToString();
    }

    private static string BuildXml() {
        var builder = new StringBuilder("<?xml version=\"1.0\" encoding=\"utf-8\"?><records>");
        for (int row = 1; row <= 1200; row++) {
            builder.Append("<record id=\"").Append(row).Append("\"><name>Item ")
                .Append(row).Append("</name><value>").Append(row * 3).Append("</value></record>");
        }
        return builder.Append("</records>").ToString();
    }

    private static string BuildYaml() {
        var builder = new StringBuilder("records:\n");
        for (int row = 1; row <= 1200; row++) {
            builder.Append("  - id: ").Append(row).Append("\n    name: Item ").Append(row)
                .Append("\n    active: ").Append(row % 2 == 0 ? "true" : "false").AppendLine();
        }
        return builder.ToString();
    }

    private static string BuildHtml() {
        var builder = new StringBuilder("<!doctype html><html><head><title>Reader benchmark</title></head><body><main>");
        for (int section = 1; section <= 160; section++) {
            builder.Append("<section><h2>Section ").Append(section).Append("</h2><p>Representative HTML content for Reader extraction and normalized chunk output.</p>")
                .Append("<ul><li>Detection</li><li>Parsing</li><li>Diagnostics</li></ul></section>");
        }
        return builder.Append("</main></body></html>").ToString();
    }

    private static string BuildRtf() {
        var builder = new StringBuilder(@"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\fs22 ");
        for (int section = 1; section <= 300; section++) {
            builder.Append(@"\b Section ").Append(section).Append(@"\b0\par ")
                .Append("Representative RTF content for Reader extraction, tables, and diagnostics.")
                .Append(@"\par ");
        }
        return builder.Append('}').ToString();
    }

    private static byte[] BuildWord() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            for (int section = 1; section <= 80; section++) {
                document.AddParagraph("Section " + section).Style = WordParagraphStyles.Heading1;
                document.AddParagraph("Representative Word content for Reader extraction, citations, and normalized output.");
            }
            document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] BuildExcel() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-reader-benchmark-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.Cell(1, 1, "Id");
                sheet.Cell(1, 2, "Name");
                sheet.Cell(1, 3, "Amount");
                sheet.Cell(1, 4, "Active");
                for (int row = 2; row <= 2501; row++) {
                    sheet.Cell(row, 1, row - 1);
                    sheet.Cell(row, 2, "Item " + (row - 1));
                    sheet.Cell(row, 3, (row - 1) * 1.25m);
                    sheet.Cell(row, 4, row % 2 == 0);
                }
                document.Save();
            }
            return File.ReadAllBytes(path);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static byte[] BuildPowerPoint() {
        using var stream = new MemoryStream();
        using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
            for (int slideNumber = 1; slideNumber <= 24; slideNumber++) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTextBox("Reader benchmark slide " + slideNumber);
                slide.AddTextBox("Representative presentation content with notes and normalized extraction.");
                slide.Notes.Text = "Speaker notes for slide " + slideNumber;
            }
            presentation.Save();
        }
        return stream.ToArray();
    }

    private static byte[] BuildBinaryPowerPoint() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        for (int slideNumber = 1; slideNumber <= 24; slideNumber++) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextBox("Reader binary benchmark slide " + slideNumber);
            slide.AddTextBox("Representative PowerPoint 97-2003 content with notes and normalized extraction.");
            slide.Notes.Text = "Speaker notes for binary slide " + slideNumber;
        }
        return presentation.ToBytes(PowerPointFileFormat.Ppt);
    }

    private static byte[] BuildPdf() {
        PdfDocument document = PdfDocument.Create();
        for (int page = 1; page <= 16; page++) {
            if (page > 1) document.PageBreak();
            document.H1("Reader benchmark page " + page);
            for (int paragraph = 1; paragraph <= 8; paragraph++) {
                int currentParagraph = paragraph;
                document.Paragraph(value => value.Text(
                    "Representative PDF paragraph " + currentParagraph + " with extraction, geometry, and diagnostic content."));
            }
        }
        return document.ToBytes();
    }

    private static byte[] BuildVisio() {
        using var stream = new MemoryStream();
        VisioDocument document = VisioDocument.Create(stream);
        VisioPage page = document.AddPage("Process").Size(16, 10);
        VisioShape? previous = null;
        for (int index = 0; index < 40; index++) {
            double x = 1 + (index % 8) * 1.8;
            double y = 1 + (index / 8) * 1.7;
            VisioShape shape = page.AddRectangle(x, y, 1.3, 0.7, "Step " + (index + 1));
            if (previous != null) page.AddConnector(previous, shape, ConnectorKind.RightAngle);
            previous = shape;
        }
        document.Save();
        return stream.ToArray();
    }

    private static byte[] BuildZip() {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            for (int entryIndex = 1; entryIndex <= 20; entryIndex++) {
                WriteTextEntry(archive, "docs/entry-" + entryIndex.ToString("D2") + ".md", BuildMarkdown());
            }
        }
        return stream.ToArray();
    }

    private static byte[] BuildEpub() {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteTextEntry(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
            WriteTextEntry(archive, "META-INF/container.xml",
                "<?xml version=\"1.0\"?><container xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles>" +
                "<rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/>" +
                "</rootfiles></container>");

            var manifest = new StringBuilder();
            var spine = new StringBuilder();
            for (int chapter = 1; chapter <= 20; chapter++) {
                manifest.Append("<item id=\"ch").Append(chapter).Append("\" href=\"chapter-")
                    .Append(chapter).Append(".xhtml\" media-type=\"application/xhtml+xml\"/>");
                spine.Append("<itemref idref=\"ch").Append(chapter).Append("\"/>");
                WriteTextEntry(archive, "OEBPS/chapter-" + chapter + ".xhtml",
                    "<?xml version=\"1.0\"?><html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Chapter " + chapter +
                    "</title></head><body><h1>Chapter " + chapter +
                    "</h1><p>Representative EPUB chapter content for Reader extraction and normalized output.</p></body></html>");
            }
            WriteTextEntry(archive, "OEBPS/content.opf",
                "<?xml version=\"1.0\"?><package version=\"3.0\" xmlns=\"http://www.idpf.org/2007/opf\">" +
                "<metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>Reader benchmark</dc:title></metadata>" +
                "<manifest>" + manifest + "</manifest><spine>" + spine + "</spine></package>");
        }
        return stream.ToArray();
    }

    private static void WriteTextEntry(
        ZipArchive archive,
        string path,
        string content,
        CompressionLevel compressionLevel = CompressionLevel.Optimal) {
        ZipArchiveEntry entry = archive.CreateEntry(path, compressionLevel);
        using Stream entryStream = entry.Open();
        using var writer = new StreamWriter(entryStream, new UTF8Encoding(false), 4096, leaveOpen: false);
        writer.Write(content);
    }
}

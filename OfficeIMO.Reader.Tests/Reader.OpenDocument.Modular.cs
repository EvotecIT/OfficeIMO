using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Testing;
using OfficeIMO.Reader.OpenDocument;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public class ReaderOpenDocumentModularTests {
    [Theory]
    [InlineData("odt")]
    [InlineData("ods")]
    [InlineData("odp")]
    public void RepeatedCellsConsumeAggregateExtractionBudgetBeforeJoiningText(string kind) {
        OdfDocument document;
        if (kind == "ods") {
            OdsDocument spreadsheet = OdsDocument.Create();
            spreadsheet.AddSheet("Data").Cell(0, 0).SetString(new string('x', 600));
            document = spreadsheet;
        } else if (kind == "odp") {
            OdpPresentation presentation = OdpPresentation.Create();
            presentation.AddSlide("Data").AddTable(OdfRect.FromCentimeters(1, 1, 10, 5), 1, 1).Cell(0, 0).Text = new string('x', 600);
            document = presentation;
        } else {
            OdtDocument text = OdtDocument.Create();
            text.AddTable(1, 1).Cell(0, 0).Text = new string('x', 600);
            document = text;
        }
        XDocument flat = document.ToFlatXml();
        flat.Descendants(OdfNamespaces.Table + "table-row").Single().SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", 200);
        flat.Descendants(OdfNamespaces.Table + "table-cell").Single().SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 256);
        using var stream = new MemoryStream();
        flat.Save(stream);
        stream.Position = 0;
        byte[] bytes = OdfDocument.LoadFlatXml(stream).ToBytes();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOpenDocumentHandler(
            new ReaderOpenDocumentOptions { MaxExtractedCharacters = 1000 }).Build();
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => reader.Read(bytes, "repeated." + kind).ToArray());
        Assert.Contains("MaxExtractedCharacters", error.Message);
    }

    [Fact]
    public void ParagraphCharacterChunksPreserveSurrogatePairsAtTheBoundary() {
        string payload = new string('x', 299) + "\U0001F600" + new string('y', 400);
        OdtDocument document = OdtDocument.Create();
        document.AddParagraph(payload);
        ReaderChunk[] chunks = CreateReader().Read(document.ToBytes(), "unicode.odt",
            new ReaderOptions { MaxChars = 300 }).ToArray();
        Assert.Equal(payload, string.Concat(chunks.Select(chunk => chunk.Text)));
        Assert.All(chunks, chunk => {
            Assert.InRange(chunk.Text.Length, 1, 300);
            Assert.False(char.IsHighSurrogate(chunk.Text[chunk.Text.Length - 1]));
            Assert.False(char.IsLowSurrogate(chunk.Text[0]));
        });
    }

    [Fact]
    public void RegisteredAdapterUsesOpenPasswordForEncryptedPackages() {
        OdtDocument document = OdtDocument.Create();
        document.AddParagraph("Encrypted reader content");
        byte[] bytes = document.ToBytes(new OdfSaveOptions {
            Encryption = new OdfEncryptionOptions { Password = "reader-test-password" }
        });
        ReaderChunk chunk = Assert.Single(CreateReader().Read(bytes, "encrypted.odt",
            new ReaderOptions { OpenPassword = "reader-test-password" }));
        Assert.Equal("Encrypted reader content", chunk.Text);
        Assert.Throws<OdfEncryptedPackageException>(() => CreateReader().Read(bytes, "encrypted.odt",
            new ReaderOptions { OpenPassword = "wrong" }).ToArray());
    }

    [Theory]
    [InlineData("ods")]
    [InlineData("odp")]
    [InlineData("odt")]
    public void CharacterBudgetSplitsLongBlocksWithoutLosingTextOrTableMetadata(string kind) {
        string payload = new string('x', 600);
        OdfDocument document;
        if (kind == "ods") {
            OdsDocument sheetDocument = OdsDocument.Create();
            OdsSheet sheet = sheetDocument.AddSheet("Data");
            sheet.Cell(0, 0).SetString("Header");
            sheet.Cell(1, 0).SetString(payload);
            document = sheetDocument;
        } else if (kind == "odp") {
            OdpPresentation slides = OdpPresentation.Create();
            slides.AddSlide("Title").AddTextBox(OdfRect.FromCentimeters(1, 1, 20, 3), payload);
            document = slides;
        } else {
            OdtDocument text = OdtDocument.Create();
            text.AddTable(1, 1).Cell(0, 0).Text = payload;
            document = text;
        }
        OfficeDocumentReader reader = CreateReader();
        ReaderChunk[] chunks = reader.Read(document.ToBytes(), "bounded." + kind,
            new ReaderOptions { MaxChars = 300 }).ToArray();
        Assert.True(chunks.Length > 1);
        Assert.All(chunks, chunk => {
            Assert.InRange(chunk.Text?.Length ?? 0, 0, 300);
            Assert.InRange(chunk.Markdown?.Length ?? 0, 0, 300);
        });
        Assert.Equal(600, string.Concat(chunks.Select(chunk => chunk.Text)).Count(character => character == 'x'));
        if (kind != "odp") Assert.Single(chunks.SelectMany(chunk => chunk.Tables ?? Array.Empty<ReaderTable>()));
    }

    [Fact]
    public void RegisteredAdapterClampsImportedHeadingLevels() {
        OdtDocument document = OdtDocument.Create();
        document.AddHeading("Imported heading", 1);
        byte[] package = RewriteHeadingLevel(document.ToBytes(), "11");

        OfficeDocumentReader reader = CreateReader();
            ReaderChunk chunk = Assert.Single(reader.Read(package, "heading.odt"));

            Assert.Equal("Imported heading", chunk.Text);
            Assert.Equal("###### Imported heading", chunk.Markdown);
            Assert.Equal("Imported heading", chunk.Location.HeadingPath);

    }

    [Fact]
    public void RegisteredAdapterHonorsRequestedOdsRange() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        sheet.Cell(0, 0).SetString("A");
        sheet.Cell(0, 1).SetString("B");
        sheet.Cell(0, 2).SetString("C");
        sheet.Cell(1, 0).SetString("Outside");
        sheet.Cell(1, 1).SetString("Two");
        sheet.Cell(1, 2).SetString("Three");
        sheet.Cell(2, 1).SetString("Outside row");

        OfficeDocumentReader reader = OfficeIMO.Reader.Tests.ReaderTestReaders.OpenDocument(a1Range: "B1:C2");
            ReaderChunk chunk = Assert.Single(reader.Read(document.ToBytes(), "range.ods"));

            Assert.Equal("B1:C2", chunk.Location.A1Range);
            ReaderTable table = Assert.Single(chunk.Tables!);
            Assert.Equal(new[] { "B", "C" }, table.Columns);
            Assert.Equal(new[] { "Two", "Three" }, Assert.Single(table.Rows));
            Assert.DoesNotContain("Outside", chunk.Text, StringComparison.Ordinal);

    }

    [Fact]
    public void RegisteredAdapterAcceptsBoundedOdsSelectionBeyondExcelRowLimit() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("LargeGrid");
        sheet.Cell(1_048_576, 0).SetString("Header");
        sheet.Cell(1_048_577, 0).SetString("Value");
        OfficeDocumentReader reader = OfficeIMO.Reader.Tests.ReaderTestReaders.OpenDocument(
            a1Range: "A1048577:A1048578");

        ReaderChunk chunk = Assert.Single(reader.Read(document.ToBytes(), "large-grid.ods"));

        Assert.Equal("A1048577:A1048578", chunk.Location.A1Range);
        ReaderTable table = Assert.Single(chunk.Tables!);
        Assert.Equal("Header", Assert.Single(table.Columns));
        Assert.Equal("Value", Assert.Single(Assert.Single(table.Rows)));
    }

    [Theory]
    [InlineData("Data!A1:B2")]
    [InlineData("A1:B2:C3")]
    [InlineData("A1:2")]
    [InlineData("B2:A1")]
    public void RegisteredAdapterRejectsNonCellOrMalformedOdsRanges(string requestedRange) {
        OdsDocument document = OdsDocument.Create();
        document.AddSheet("Data").Cell(0, 0).SetString("A");
        OfficeDocumentReader reader = OfficeIMO.Reader.Tests.ReaderTestReaders.OpenDocument(a1Range: requestedRange);

        Assert.Throws<FormatException>(() => reader.Read(document.ToBytes(), "range.ods").ToList());
    }

    [Fact]
    public void RegisteredAdapterEmitsSlideAlignedOdpChunkWithNotesAndTable() {
        OdpPresentation document = OdpPresentation.Create();
        OdpSlide slide = document.AddSlide("Summary");
        slide.AddTextBox(OdfRect.FromCentimeters(1, 1, 20, 3), "Native presentation");
        OdpTable table = slide.AddTable(OdfRect.FromCentimeters(1, 5, 12, 4), 2, 2, "Metrics");
        table.Cell(0, 0).Text = "Name";
        table.Cell(0, 1).Text = "Value";
        table.Cell(1, 0).Text = "Revenue";
        table.Cell(1, 1).Text = "42";
        slide.GetOrCreateSpeakerNotes().AddParagraph("Explain the result.");

        OfficeDocumentReader reader = CreateReader();
            ReaderChunk chunk = Assert.Single(reader.Read(document.ToBytes(), "summary.odp"));

            Assert.Equal(1, chunk.Location.Slide);
            Assert.Equal("Summary", chunk.Location.HeadingPath);
            Assert.Contains("Native presentation", chunk.Text, StringComparison.Ordinal);
            Assert.Contains("Notes: Explain the result.", chunk.Text, StringComparison.Ordinal);
            Assert.Equal("42", Assert.Single(chunk.Tables!).Rows[1][1]);

    }

    [Fact]
    public void RegisteredAdapterReportsOdpColumnTruncationSeparately() {
        OdpPresentation document = OdpPresentation.Create();
        OdpSlide slide = document.AddSlide("Wide");
        OdpTable table = slide.AddTable(OdfRect.FromCentimeters(1, 1, 20, 4), 1, 257, "Wide table");
        table.Cell(0, 256).Text = "Truncated column";

        ReaderChunk chunk = Assert.Single(CreateReader().Read(document.ToBytes(), "wide.odp"));

        ReaderTable extracted = Assert.Single(chunk.Tables!);
        Assert.True(extracted.Truncated);
        Assert.Equal(256, extracted.Columns.Count);
        Assert.Contains(chunk.Warnings!, warning => warning.Contains("columns were truncated", StringComparison.Ordinal));
        Assert.DoesNotContain(chunk.Warnings!, warning => warning.Contains("rows were truncated", StringComparison.Ordinal));
    }

    [Fact]
    public void RegisteredAdapterEmitsBoundedOdsSheetTableChunk() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Metrics");
        sheet.Cell(0, 0).SetString("Name");
        sheet.Cell(0, 1).SetString("Value");
        sheet.Cell(1, 0).SetString("Revenue");
        sheet.Cell(1, 1).SetDecimal(42.5m);

        OfficeDocumentReader reader = CreateReader();
            ReaderChunk chunk = Assert.Single(reader.Read(document.ToBytes(), "metrics.ods"));

            Assert.Equal("Metrics", chunk.Location.Sheet);
            Assert.Equal("A1:B2", chunk.Location.A1Range);
            ReaderTable table = Assert.Single(chunk.Tables!);
            Assert.Equal(new[] { "Name", "Value" }, table.Columns);
            Assert.Equal("Revenue", table.Rows[0][0]);
            Assert.Equal("42.5", table.Rows[0][1]);

    }

    [Fact]
    public void RegisteredAdapterEmitsOdtHeadingParagraphAndTableChunks() {
        OdtDocument document = OdtDocument.Create();
        document.AddHeading("Policy", 1);
        document.AddParagraph("Native OpenDocument text.");
        OdtTable table = document.AddTable(2, 2, "Approvals");
        table.Cell(0, 0).Text = "Owner";
        table.Cell(0, 1).Text = "Status";
        table.Cell(1, 0).Text = "Operations";
        table.Cell(1, 1).Text = "Approved";

        OfficeDocumentReader reader = CreateReader();
            IReadOnlyList<ReaderChunk> chunks = reader.Read(document.ToBytes(), "policy.odt").ToList();

            Assert.Equal(ReaderInputKind.OpenDocument, reader.DetectKind("policy.odt"));
            Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "heading" && chunk.Text == "Policy");
            Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "paragraph" && chunk.Location.HeadingPath == "Policy");
            ReaderChunk tableChunk = Assert.Single(chunks, chunk => chunk.Location.SourceBlockKind == "table");
            ReaderTable extracted = Assert.Single(tableChunk.Tables!);
            Assert.Equal("Approvals", extracted.Title);
            Assert.Equal("Approved", extracted.Rows[1][1]);
            Assert.All(chunks, chunk => Assert.Equal(ReaderInputKind.OpenDocument, chunk.Kind));

    }

    [Fact]
    public void RegisteredAdapterReportsOdtColumnTruncationSeparately() {
        OdtDocument document = OdtDocument.Create();
        OdtTable table = document.AddTable(1, 257, "Wide");
        table.Cell(0, 256).Text = "Truncated column";

        ReaderChunk chunk = Assert.Single(CreateReader().Read(document.ToBytes(), "wide.odt"));

        ReaderTable extracted = Assert.Single(chunk.Tables!);
        Assert.True(extracted.Truncated);
        Assert.Equal(256, extracted.Columns.Count);
        Assert.Contains(chunk.Warnings!, warning => warning.Contains("columns were truncated", StringComparison.Ordinal));
        Assert.DoesNotContain(chunk.Warnings!, warning => warning.Contains("rows were truncated", StringComparison.Ordinal));
    }

    private static OfficeDocumentReader CreateReader() {
        return new OfficeDocumentReaderBuilder().AddOpenDocumentHandler().Build();
    }

    private static byte[] RewriteHeadingLevel(byte[] package, string level) {
        return OdfTestPackageRewriter.Rewrite(package, (name, bytes) => {
            if (name == "content.xml") {
                XDocument content = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
                content.Descendants(text + "h").Single().SetAttributeValue(text + "outline-level", level);
                return Encoding.UTF8.GetBytes(content.ToString(SaveOptions.DisableFormatting));
            }
            return bytes;
        });
    }
}

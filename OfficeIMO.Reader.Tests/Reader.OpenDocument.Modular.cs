using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Testing;
using OfficeIMO.Reader.OpenDocument;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public class ReaderOpenDocumentModularTests {
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

using OfficeIMO.OpenDocument;
using OfficeIMO.Reader.OpenDocument;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public class ReaderOpenDocumentModularTests {
    [Fact]
    public void RegisteredAdapterEmitsSlideAlignedOdpChunkWithNotesAndTable() {
        using OdpPresentation document = OdpPresentation.Create();
        OdpSlide slide = document.AddSlide("Summary");
        slide.AddTextBox(OdfRect.FromCentimeters(1, 1, 20, 3), "Native presentation");
        OdpTable table = slide.AddTable(OdfRect.FromCentimeters(1, 5, 12, 4), 2, 2, "Metrics");
        table.Cell(0, 0).Text = "Name";
        table.Cell(0, 1).Text = "Value";
        table.Cell(1, 0).Text = "Revenue";
        table.Cell(1, 1).Text = "42";
        slide.GetOrCreateSpeakerNotes().AddParagraph("Explain the result.");

        DocumentReaderOpenDocumentRegistrationExtensions.RegisterOpenDocumentHandler(replaceExisting: true);
        try {
            ReaderChunk chunk = Assert.Single(DocumentReader.Read(document.ToBytes(), "summary.odp"));

            Assert.Equal(1, chunk.Location.Slide);
            Assert.Equal("Summary", chunk.Location.HeadingPath);
            Assert.Contains("Native presentation", chunk.Text, StringComparison.Ordinal);
            Assert.Contains("Notes: Explain the result.", chunk.Text, StringComparison.Ordinal);
            Assert.Equal("42", Assert.Single(chunk.Tables!).Rows[1][1]);
        } finally {
            DocumentReaderOpenDocumentRegistrationExtensions.UnregisterOpenDocumentHandler();
        }
    }

    [Fact]
    public void RegisteredAdapterEmitsBoundedOdsSheetTableChunk() {
        using OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Metrics");
        sheet.Cell(0, 0).SetString("Name");
        sheet.Cell(0, 1).SetString("Value");
        sheet.Cell(1, 0).SetString("Revenue");
        sheet.Cell(1, 1).SetDecimal(42.5m);

        DocumentReaderOpenDocumentRegistrationExtensions.RegisterOpenDocumentHandler(replaceExisting: true);
        try {
            ReaderChunk chunk = Assert.Single(DocumentReader.Read(document.ToBytes(), "metrics.ods"));

            Assert.Equal("Metrics", chunk.Location.Sheet);
            Assert.Equal("A1:B2", chunk.Location.A1Range);
            ReaderTable table = Assert.Single(chunk.Tables!);
            Assert.Equal(new[] { "Name", "Value" }, table.Columns);
            Assert.Equal("Revenue", table.Rows[0][0]);
            Assert.Equal("42.5", table.Rows[0][1]);
        } finally {
            DocumentReaderOpenDocumentRegistrationExtensions.UnregisterOpenDocumentHandler();
        }
    }

    [Fact]
    public void RegisteredAdapterEmitsOdtHeadingParagraphAndTableChunks() {
        using OdtDocument document = OdtDocument.Create();
        document.AddHeading("Policy", 1);
        document.AddParagraph("Native OpenDocument text.");
        OdtTable table = document.AddTable(2, 2, "Approvals");
        table.Cell(0, 0).Text = "Owner";
        table.Cell(0, 1).Text = "Status";
        table.Cell(1, 0).Text = "Operations";
        table.Cell(1, 1).Text = "Approved";

        DocumentReaderOpenDocumentRegistrationExtensions.RegisterOpenDocumentHandler(replaceExisting: true);
        try {
            IReadOnlyList<ReaderChunk> chunks = DocumentReader.Read(document.ToBytes(), "policy.odt").ToList();

            Assert.Equal(ReaderInputKind.OpenDocument, DocumentReader.DetectKind("policy.odt"));
            Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "heading" && chunk.Text == "Policy");
            Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "paragraph" && chunk.Location.HeadingPath == "Policy");
            ReaderChunk tableChunk = Assert.Single(chunks, chunk => chunk.Location.SourceBlockKind == "table");
            ReaderTable extracted = Assert.Single(tableChunk.Tables!);
            Assert.Equal("Approvals", extracted.Title);
            Assert.Equal("Approved", extracted.Rows[1][1]);
            Assert.All(chunks, chunk => Assert.Equal(ReaderInputKind.OpenDocument, chunk.Kind));
        } finally {
            DocumentReaderOpenDocumentRegistrationExtensions.UnregisterOpenDocumentHandler();
        }
    }
}

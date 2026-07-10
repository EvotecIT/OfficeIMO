using OfficeIMO.OpenDocument;
using OfficeIMO.Reader.OpenDocument;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public class ReaderOpenDocumentModularTests {
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

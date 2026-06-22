using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfIncrementalUpdaterTests {
    [Fact]
    public void UpdateMetadata_AppendsIncrementalRevisionAndPreservesContent() {
        byte[] original = PdfDocument.Create()
            .Meta(title: "Original title", author: "Original author")
            .Paragraph(paragraph => paragraph.Text("Incremental update body text"))
            .ToBytes();

        byte[] updated = PdfIncrementalUpdater.UpdateMetadata(original, title: "Updated title");
        PdfDocumentInfo info = PdfInspector.Inspect(updated);

        Assert.True(updated.Length > original.Length);
        Assert.Equal("Updated title", info.Metadata.Title);
        Assert.Equal("Original author", info.Metadata.Author);
        Assert.Contains("Incremental update body text", PdfTextExtractor.ExtractAllText(updated), StringComparison.Ordinal);
        Assert.True(info.Security.HasIncrementalUpdates);
        Assert.True(info.Security.HasPreviousRevision);
        Assert.True(info.Security.RevisionCount >= 2);
        Assert.Contains(info.Security.Revisions, revision => revision.HasPreviousRevision);
    }

    [Fact]
    public void UpdateMetadata_PreservesNonZeroRootGenerationInAppendedTrailer() {
        byte[] original = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 2 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] >>",
            "endobj",
            "trailer",
            "<< /Root 1 2 R /Size 4 >>",
            "startxref",
            "123",
            "%%EOF"
        }));

        byte[] updated = PdfIncrementalUpdater.UpdateMetadata(original, title: "Updated title");
        string raw = PdfEncoding.Latin1GetString(updated);

        Assert.Contains("/Root 1 2 R", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/Root 1 0 R /Info", raw, StringComparison.Ordinal);
        Assert.Equal("Updated title", PdfInspector.Inspect(updated).Metadata.Title);
    }
}

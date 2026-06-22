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
}

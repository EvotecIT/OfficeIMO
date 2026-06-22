using System.Globalization;
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
    public void UpdateMetadata_PreservesCatalogReferenceGenerationInAppendedTrailer() {
        byte[] original = BuildMetadataPdfWithCatalogGeneration();

        byte[] updated = PdfIncrementalUpdater.UpdateMetadata(original, title: "Updated title");

        string updatedText = PdfEncoding.Latin1GetString(updated);
        int appendedTrailer = updatedText.LastIndexOf("trailer", StringComparison.Ordinal);
        Assert.True(appendedTrailer >= PdfEncoding.Latin1GetString(original).Length);
        Assert.Contains("/Root 1 2 R", updatedText.Substring(appendedTrailer), StringComparison.Ordinal);
        Assert.Equal("Updated title", PdfInspector.Inspect(updated).Metadata.Title);
    }

    private static byte[] BuildMetadataPdfWithCatalogGeneration() {
        var entries = new List<(int ObjectNumber, int Generation, string Body)> {
            (1, 2, "<< /Type /Catalog /Pages 2 0 R >>"),
            (2, 0, "<< /Type /Pages /Count 1 /Kids [3 0 R] >>"),
            (3, 0, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>"),
            (4, 0, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"),
            (5, 0, BuildStream(Encoding.ASCII.GetBytes("BT /F1 12 Tf 72 720 Td (Metadata generation) Tj ET")))
        };

        var builder = new StringBuilder();
        builder.AppendLine("%PDF-1.7");
        foreach ((int objectNumber, int generation, string body) in entries) {
            builder.Append(objectNumber.ToString(CultureInfo.InvariantCulture)).Append(' ')
                .Append(generation.ToString(CultureInfo.InvariantCulture)).AppendLine(" obj");
            builder.AppendLine(body);
            builder.AppendLine("endobj");
        }

        builder.AppendLine("trailer");
        builder.AppendLine("<< /Root 1 2 R /Size 6 >>");
        builder.AppendLine("startxref");
        builder.AppendLine("123");
        builder.AppendLine("%%EOF");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }

    private static string BuildStream(byte[] data) =>
        "<< /Length " + data.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" +
        Encoding.ASCII.GetString(data) +
        "\nendstream";
}

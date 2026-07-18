using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfIncrementalUpdaterTests {
    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4)]
    public void UpdateMetadata_AppendsEncryptedRevisionWithOwnerAuthorization(PdfStandardEncryptionAlgorithm algorithm) {
        var encryption = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            Algorithm = algorithm
        };
        byte[] original = PdfDocument.Create(new PdfOptions { IncludeXmpMetadata = true }.SetEncryption(encryption))
            .Meta(title: "Encrypted original", author: "Encrypted author")
            .Paragraph(paragraph => paragraph.Text("Encrypted incremental body"))
            .ToBytes();
        var ownerOptions = new PdfReadOptions { Password = "owner" };
        var userOptions = new PdfReadOptions { Password = "open" };
        PdfDocumentSecurityInfo before = PdfSyntax.ReadDocumentSecurityInfo(original, ownerOptions);

        byte[] updated = PdfIncrementalUpdater.UpdateMetadata(
            original,
            title: "Encrypted updated",
            readOptions: ownerOptions);

        Assert.True(updated.AsSpan(0, original.Length).SequenceEqual(original));
        Assert.Throws<PdfPasswordRequiredException>(() => PdfInspector.Inspect(updated));
        Assert.Throws<PdfInvalidPasswordException>(() => PdfInspector.Inspect(updated, new PdfReadOptions { Password = "wrong" }));
        PdfDocumentInfo userInfo = PdfInspector.Inspect(updated, userOptions);
        PdfDocumentSecurityInfo after = PdfSyntax.ReadDocumentSecurityInfo(updated, ownerOptions);
        Assert.Equal("Encrypted updated", userInfo.Metadata.Title);
        Assert.Equal("Encrypted updated", userInfo.XmpMetadata!.Title);
        Assert.Equal("Encrypted author", userInfo.Metadata.Author);
        Assert.Contains(
            "Encrypted incremental body",
            PdfTextExtractor.ExtractAllText(updated, (PdfTextLayoutOptions?)null, userOptions),
            StringComparison.Ordinal);
        Assert.True(after.HasEncryption);
        Assert.Equal(before.EncryptObjectNumber, after.EncryptObjectNumber);
        Assert.True(after.HasTrailerId);
        Assert.True(after.HasIncrementalUpdates);
        Assert.Equal(before.LastStartXrefOffset, after.Revisions[after.Revisions.Count - 1].PreviousXrefOffset);
        string appended = PdfEncoding.Latin1GetString(updated, original.Length, updated.Length - original.Length);
        Assert.Contains("/Encrypt " + before.EncryptObjectNumber + " 0 R", appended, StringComparison.Ordinal);
        Assert.Contains("/ID [", appended, StringComparison.Ordinal);
    }

    [Fact]
    public void UpdateMetadata_SynchronizesXmpAndPreservesExtensionSchemas() {
        var options = new PdfOptions { IncludeXmpMetadata = true }
            .SetElectronicInvoiceMetadata(new PdfElectronicInvoiceMetadata("ORDER", "invoice.xml", "1.0", "EN 16931"));
        byte[] original = PdfDocument.Create(options)
            .Meta(title: "Original XMP", author: "Original author", keywords: "first; second")
            .Paragraph(paragraph => paragraph.Text("XMP extension preservation"))
            .ToBytes();

        byte[] updated = PdfIncrementalUpdater.UpdateMetadata(
            original,
            title: "Synchronized XMP",
            author: string.Empty,
            keywords: "third, fourth");
        PdfDocumentInfo info = PdfInspector.Inspect(updated);

        Assert.True(updated.AsSpan(0, original.Length).SequenceEqual(original));
        Assert.Equal("Synchronized XMP", info.Metadata.Title);
        Assert.Null(info.Metadata.Author);
        Assert.Equal("third, fourth", info.Metadata.Keywords);
        Assert.Equal("Synchronized XMP", info.XmpMetadata!.Title);
        Assert.Null(info.XmpMetadata.Creator);
        Assert.Equal(new[] { "third", "fourth" }, info.XmpMetadata.Subjects);
        Assert.Equal("ORDER", info.XmpMetadata.ElectronicInvoiceDocumentType);
        Assert.Equal("invoice.xml", info.XmpMetadata.ElectronicInvoiceDocumentFileName);
        Assert.Equal("EN 16931", info.XmpMetadata.ElectronicInvoiceConformanceLevel);
    }

    [Fact]
    public void AppendMetadataRevision_CanCreateSynchronizedXmpPacket() {
        byte[] original = PdfDocument.Create()
            .Meta(title: "Info only")
            .Paragraph(paragraph => paragraph.Text("Create XMP append"))
            .ToBytes();

        PdfDocument updated = PdfDocument.Open(original)
            .AppendMetadataRevision(title: "Created XMP", keywords: "one;two", createXmpMetadata: true);
        PdfDocumentInfo info = updated.Inspect();

        Assert.Equal("Created XMP", info.Metadata.Title);
        Assert.Equal("Created XMP", info.XmpMetadata!.Title);
        Assert.Equal(new[] { "one", "two" }, info.XmpMetadata.Subjects);
    }

    [Fact]
    public void FluentAppendMetadataRevision_ReusesEncryptedDocumentReadOptions() {
        byte[] original = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Meta(title: "Before fluent append")
            .Paragraph(paragraph => paragraph.Text("Fluent encrypted body"))
            .ToBytes();
        var ownerOptions = new PdfReadOptions { Password = "owner" };

        PdfDocument updated = PdfDocument.Open(original, ownerOptions)
            .AppendMetadataRevision(title: "After fluent append");

        Assert.Equal("After fluent append", updated.Read.Metadata().Title);
        Assert.Contains("Fluent encrypted body", updated.Read.Text(), StringComparison.Ordinal);
    }

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

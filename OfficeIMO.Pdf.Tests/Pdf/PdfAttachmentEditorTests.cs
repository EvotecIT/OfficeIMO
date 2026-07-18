using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAttachmentEditorTests {
    [Fact]
    public void Edit_AddsReplacesRenamesRemovesAndValidatesAttachmentMetadata() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.Text("Attachment edit proof"))
            .AttachFile("alpha.txt", Encoding.UTF8.GetBytes("alpha"), "text/plain")
            .AttachFile("beta.txt", Encoding.UTF8.GetBytes("old beta"), "text/plain")
            .AttachFile("obsolete.bin", new byte[] { 1, 2, 3 }, "application/octet-stream")
            .ToBytes();
        var created = new DateTimeOffset(2026, 7, 11, 12, 30, 45, TimeSpan.FromHours(2));
        var modified = created.AddHours(1);

        PdfAttachmentEditResult result = PdfDocument.Open(source).Attachments.Edit(attachments => attachments
            .Rename("alpha.txt", "renamed.txt")
            .Replace("beta.txt", new PdfEmbeddedFile("beta.txt", Encoding.UTF8.GetBytes("new beta"), "text/plain", description: "replacement"))
            .Remove("obsolete.bin")
            .Add(new PdfEmbeddedFile("data.xml", Encoding.UTF8.GetBytes("<data />"), "application/xml", PdfAssociatedFileRelationship.Data, "associated data", created, modified)));

        Assert.Equal(PdfMutationExecutionMode.FullRewrite, result.MutationPlan.ExecutionMode);
        Assert.Contains(PdfMutationProof.AttachmentReadback, result.MutationPlan.RequiredProofs);
        Assert.True(result.PreservationReport.IsPreserved, string.Join(" ", result.PreservationReport.Issues.Select(static issue => issue.Message)));
        Assert.True(result.IsValid);
        Assert.Equal(3, result.Validations.Count);
        Assert.All(result.Validations, validation => { Assert.True(validation.IsValid); Assert.Equal(32, validation.Checksum.Length); });

        byte[] output = result.ToBytes();
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(output);
        Assert.Equal(new[] { "beta.txt", "data.xml", "renamed.txt" }, attachments.Select(static attachment => attachment.FileName));
        PdfExtractedAttachment data = Assert.Single(attachments, static attachment => attachment.FileName == "data.xml");
        Assert.Equal(PdfAssociatedFileRelationship.Data, data.Relationship);
        Assert.Equal("application/xml", data.MimeType);
        Assert.Equal(created, data.CreationDate);
        Assert.Equal(modified, data.ModificationDate);
        Assert.Contains("/AF [", Encoding.ASCII.GetString(output), StringComparison.Ordinal);
        Assert.Equal("new beta", Encoding.UTF8.GetString(Assert.Single(attachments, static attachment => attachment.FileName == "beta.txt").Bytes));
        Assert.Single(PdfAttachmentExtractor.ExtractAttachmentsByFileName(output, "renamed.txt"));
        Assert.Single(PdfAttachmentExtractor.ExtractAttachments(output, static attachment => attachment.Relationship == PdfAssociatedFileRelationship.Data));
    }

    [Fact]
    public void Edit_RejectsMissingAndDuplicateNames() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Attachment conflicts")).AttachFile("one.txt", new byte[] { 1 }).ToBytes();
        Assert.Throws<KeyNotFoundException>(() => PdfAttachmentEditor.Remove(source, "missing.txt"));
        Assert.Throws<ArgumentException>(() => PdfAttachmentEditor.Add(source, new PdfEmbeddedFile("one.txt", new byte[] { 2 })));
    }

    [Fact]
    public void Edit_RemovePrunesPageAssociatedFilePayloads() {
        byte[] source = PdfAssociatedFileTestSupport.BuildPageAssociatedFilePdf();
        Assert.Single(PdfAttachmentExtractor.ExtractAttachments(source));

        byte[] output = PdfAttachmentEditor.Remove(source, "page.txt").ToBytes();

        Assert.Empty(PdfAttachmentExtractor.ExtractAttachments(output));
        string raw = Encoding.ASCII.GetString(output);
        Assert.DoesNotContain("/AF", raw, StringComparison.Ordinal);
        Assert.DoesNotContain(PdfAssociatedFileTestSupport.Payload, raw, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadDocument_AppliesDecodedStreamLimitToPageAssociatedFiles() {
        byte[] source = PdfAssociatedFileTestSupport.BuildPageAssociatedFilePdf();
        var options = new PdfReadOptions { Limits = new PdfReadLimits { MaxDecodedStreamBytes = 8 } };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(source, options));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(8, exception.Limit);
    }
}

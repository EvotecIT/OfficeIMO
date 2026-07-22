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
    public void Edit_RemoveDeletesUnreferencedEmbeddedFileObjectsFromTheRewrittenGraph() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.Text("Attachment graph cleanup"))
            .AttachFile("hidden-marker.txt", Encoding.ASCII.GetBytes("hidden-attachment-marker"), "text/plain")
            .ToBytes();

        byte[] output = PdfAttachmentEditor.Remove(source, "hidden-marker.txt").ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(output);

        Assert.Empty(PdfAttachmentExtractor.ExtractAttachments(output));
        Assert.DoesNotContain(objects.Values, static item =>
            item.Value is PdfStream stream &&
            string.Equals(stream.Dictionary.Get<PdfName>("Type")?.Name, "EmbeddedFile", StringComparison.Ordinal));
        Assert.DoesNotContain(objects.Values, static item =>
            (item.Value as PdfDictionary)?.Items.ContainsKey("EF") == true);
        Assert.DoesNotContain("hidden-attachment-marker", PdfEncoding.Latin1GetString(output), StringComparison.Ordinal);
    }

    [Fact]
    public void Edit_RetainsAndReconnectsFileAttachmentAnnotations() {
        byte[] source = PdfAssociatedFileTestSupport.BuildFileAttachmentAnnotationPdf();

        byte[] output = PdfAttachmentEditor.Rename(source, "page.txt", "renamed.txt").ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(output);
        PdfDictionary annotation = Assert.Single(
            objects.Values.Select(static item => item.Value as PdfDictionary),
            static dictionary => string.Equals(dictionary?.Get<PdfName>("Subtype")?.Name, "FileAttachment", StringComparison.Ordinal))!;
        PdfReference fileSpecificationReference = Assert.IsType<PdfReference>(annotation.Items["FS"]);
        PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[fileSpecificationReference.ObjectNumber].Value);

        Assert.Equal("renamed.txt", fileSpecification.Get<PdfStringObj>("UF")?.Value);
        Assert.Single(PdfAttachmentExtractor.ExtractAttachmentsByFileName(output, "renamed.txt"));
    }

    [Fact]
    public void Edit_RetainsStandaloneFileAttachmentAnnotationWhenAddingAnotherAttachment() {
        byte[] source = PdfAssociatedFileTestSupport.BuildStandaloneFileAttachmentAnnotationPdf();
        Assert.Single(PdfAttachmentExtractor.ExtractAttachmentsByFileName(source, "page.txt"));

        byte[] output = PdfAttachmentEditor.Add(
            source,
            new PdfEmbeddedFile("new.txt", Encoding.UTF8.GetBytes("new payload"), "text/plain")).ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(output);

        Assert.Single(objects.Values, static item =>
            (item.Value as PdfDictionary)?.Get<PdfName>("Subtype")?.Name == "FileAttachment");
        Assert.Single(PdfAttachmentExtractor.ExtractAttachmentsByFileName(output, "page.txt"));
        Assert.Single(PdfAttachmentExtractor.ExtractAttachmentsByFileName(output, "new.txt"));
    }

    [Fact]
    public void EditSession_AllowsDuplicateNamesAlreadyPresentInTheSource() {
        var session = new PdfAttachmentEditSession(new[] {
            new PdfEmbeddedFile("duplicate.txt", new byte[] { 1 }),
            new PdfEmbeddedFile("duplicate.txt", new byte[] { 2 })
        });

        Assert.Equal(2, session.Attachments.Count);
    }

    [Fact]
    public void Edit_PreservesDuplicateNamedAnnotationPayloadIdentity() {
        byte[] source = PdfAssociatedFileTestSupport.BuildDuplicateNamedFileAttachmentAnnotationsPdf();

        byte[] output = PdfAttachmentEditor.Add(
            source,
            new PdfEmbeddedFile("new.txt", Encoding.UTF8.GetBytes("new payload"), "text/plain"))
            .ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(output);
        PdfDictionary[] annotations = objects.Values
            .Select(static item => item.Value as PdfDictionary)
            .Where(static dictionary =>
                string.Equals(dictionary?.Get<PdfName>("Subtype")?.Name,
                    "FileAttachment", StringComparison.Ordinal))
            .Select(static dictionary => dictionary!)
            .OrderBy(static dictionary => Assert.IsType<PdfNumber>(
                Assert.IsType<PdfArray>(dictionary.Items["Rect"]).Items[0]).Value)
            .ToArray();

        Assert.Equal(2, annotations.Length);
        Assert.Equal("FIRST-DUPLICATE-PAYLOAD", ReadAnnotationPayload(objects, annotations[0]));
        Assert.Equal("SECOND-DUPLICATE-PAYLOAD", ReadAnnotationPayload(objects, annotations[1]));
        Assert.Equal(2, PdfAttachmentExtractor.ExtractAttachmentsByFileName(output, "duplicate.txt").Count);
    }

    [Fact]
    public void Edit_RemovesFileAttachmentAnnotationWhenItsPayloadIsRemoved() {
        byte[] source = PdfAssociatedFileTestSupport.BuildFileAttachmentAnnotationPdf();

        byte[] output = PdfAttachmentEditor.Remove(source, "page.txt").ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(output);

        Assert.DoesNotContain(objects.Values, static item =>
            (item.Value as PdfDictionary)?.Get<PdfName>("Subtype")?.Name == "FileAttachment");
    }

    [Fact]
    public void ReadDocument_AppliesDecodedStreamLimitToPageAssociatedFiles() {
        byte[] source = PdfAssociatedFileTestSupport.BuildPageAssociatedFilePdf();
        var options = new PdfReadOptions { Limits = new PdfReadLimits { MaxDecodedStreamBytes = 8 } };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(source, options));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(8, exception.Limit);
    }

    private static string ReadAnnotationPayload(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary annotation) {
        PdfReference fileSpecReference = Assert.IsType<PdfReference>(annotation.Items["FS"]);
        PdfDictionary fileSpec = Assert.IsType<PdfDictionary>(objects[fileSpecReference.ObjectNumber].Value);
        PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(fileSpec.Items["EF"]);
        PdfReference embeddedFileReference = Assert.IsType<PdfReference>(embeddedFiles.Items["UF"]);
        PdfStream embeddedFile = Assert.IsType<PdfStream>(objects[embeddedFileReference.ObjectNumber].Value);
        return Encoding.ASCII.GetString(embeddedFile.Data);
    }
}

using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfProvenanceTests {
    [Fact]
    public void InspectAndRemoveUseTheExactC2paAssociatedFileProfile() {
        byte[] manifest = CreateManifestStore();
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("PDF provenance"))
            .AttachFile(new PdfEmbeddedFile(
                "content-credential.c2pa",
                manifest,
                "application/c2pa",
                PdfAssociatedFileRelationship.C2paManifest))
            .AttachFile(new PdfEmbeddedFile(
                "keep.txt",
                Encoding.UTF8.GetBytes("keep"),
                "text/plain",
                PdfAssociatedFileRelationship.Supplement))
            .ToBytes();

        OfficeProvenanceReport report = PdfProvenance.Inspect(pdf);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf);
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(result.ToArray());

        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence);
        Assert.True(evidence.IsStructurallyValid);
        Assert.Equal(OfficeProvenanceAssetFormat.Pdf, report.Format);
        Assert.Empty(result.After.Evidence);
        PdfExtractedAttachment retained = Assert.Single(attachments);
        Assert.Equal("keep.txt", retained.FileName);
        Assert.Equal("keep", Encoding.UTF8.GetString(retained.Bytes));
        Assert.Equal(PdfAssociatedFileRelationship.Supplement, retained.Relationship);
    }

    [Fact]
    public void MalformedC2paAssociatedFileIsPreservedByDefault() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Malformed provenance"))
            .AttachFile(new PdfEmbeddedFile(
                "content-credential.c2pa",
                Encoding.ASCII.GetBytes("not-a-manifest"),
                "application/c2pa",
                PdfAssociatedFileRelationship.C2paManifest))
            .ToBytes();

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf);

        Assert.False(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.False(result.Before.Evidence[0].IsStructurallyValid);
        Assert.Equal(pdf, result.ToArray());
    }

    [Fact]
    public void CallerCanExplicitlyRemoveMalformedCandidateCarrier() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Malformed provenance"))
            .AttachFile(new PdfEmbeddedFile(
                "content-credential.c2pa",
                Encoding.ASCII.GetBytes("not-a-manifest"),
                "application/c2pa",
                PdfAssociatedFileRelationship.C2paManifest))
            .ToBytes();
        var options = new OfficeProvenanceRemovalOptions { RequireStructurallyValidCarrier = false };

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf, options);

        Assert.True(result.WasChanged);
        Assert.Empty(PdfAttachmentExtractor.ExtractAttachments(result.ToArray()));
    }

    [Fact]
    public void InspectionEnforcesTheSharedAssetLimitBeforePdfParsing() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Bounded")).ToBytes();

        Assert.Throws<InvalidDataException>(() => PdfProvenance.Inspect(
            pdf,
            new OfficeProvenanceOptions {
                MaxAssetBytes = pdf.Length - 1,
                MaxManifestBytes = Math.Min(64, pdf.Length - 1)
            }));
    }

    private static byte[] CreateManifestStore() {
        byte[] data = new byte[38];
        WriteBigEndian(data, 0, data.Length);
        Encoding.ASCII.GetBytes("jumb").CopyTo(data, 4);
        WriteBigEndian(data, 8, 30);
        Encoding.ASCII.GetBytes("jumd").CopyTo(data, 12);
        new byte[] { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 }.CopyTo(data, 16);
        data[32] = 0x02;
        Encoding.ASCII.GetBytes("c2pa").CopyTo(data, 33);
        return data;
    }

    private static void WriteBigEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}

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
                "keep.txt",
                Encoding.UTF8.GetBytes("keep"),
                "text/plain",
                PdfAssociatedFileRelationship.Supplement))
            .AttachFile(new PdfEmbeddedFile(
                "content-credential.c2pa",
                manifest,
                "application/c2pa",
                PdfAssociatedFileRelationship.C2paManifest))
            .ToBytes();

        OfficeProvenanceReport report = PdfProvenance.Inspect(pdf);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf);
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(result.ToArray());

        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence);
        Assert.True(evidence.IsStructurallyValid);
        Assert.Equal(OfficeProvenanceAssetFormat.Pdf, report.Format);
        Assert.Empty(result.After.Evidence);
        Assert.All(result.Changes, change => Assert.Equal(0, change.RemovedBytes));
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

    [Fact]
    public void CandidateWithoutAnAssociatedFileReferenceIsNotStructurallyValid() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Detached PDF provenance"))
            .AttachFile(new PdfEmbeddedFile(
                "content-credential.c2pa",
                CreateManifestStore(),
                "application/c2pa",
                PdfAssociatedFileRelationship.C2paManifest))
            .ToBytes();
        byte[] detached = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            catalog.Items.Remove("AF");
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(detached);

        Assert.Single(result.Before.Evidence);
        Assert.False(result.Before.Evidence[0].IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(detached, result.ToArray());
    }

    [Fact]
    public void PageAssociatedFileReferenceDoesNotSatisfyTheDocumentCredentialProfile() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] pageAssociated = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfObject associatedFiles = catalog.Items["AF"];
            catalog.Items.Remove("AF");
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values
                .Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            page.Items["AF"] = associatedFiles;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(pageAssociated);

        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence);
        Assert.False(evidence.IsStructurallyValid);
    }

    [Fact]
    public void RemovalMapsDuplicateCarrierDescriptorsToDistinctAttachmentIndices() {
        byte[] duplicated = DuplicateCandidateAroundRetainedAttachment(CreatePdfWithCandidateAndRetainedAttachment(), copies: 2);

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(duplicated);
        IReadOnlyList<PdfExtractedAttachment> retained = PdfAttachmentExtractor.ExtractAttachments(result.ToArray());

        Assert.Equal(2, result.Before.Evidence.Count);
        PdfExtractedAttachment attachment = Assert.Single(retained);
        Assert.Equal("keep.txt", attachment.FileName);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void AttachmentRewriteBudgetCountsEveryProjectedDescriptorCopy() {
        byte[] duplicated = DuplicateCandidateAroundRetainedAttachment(CreatePdfWithCandidateAndRetainedAttachment(), copies: 3);
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxManifestBytes = 64;
        options.Limits.MaxExpandedContainerBytes = 64;

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfProvenance.Remove(duplicated, options));

        Assert.Equal(PdfReadLimitKind.AttachmentBytes, exception.Kind);
    }

    [Fact]
    public void InspectionDoesNotDecodeUnrelatedAttachmentsAgainstTheProvenanceBudget() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Bounded provenance"))
            .AttachFile(new PdfEmbeddedFile(
                "unrelated.bin",
                new byte[1024 * 1024],
                "application/octet-stream",
                PdfAssociatedFileRelationship.Supplement))
            .AttachFile(new PdfEmbeddedFile(
                "content-credential.c2pa",
                CreateManifestStore(),
                "application/c2pa",
                PdfAssociatedFileRelationship.C2paManifest))
            .ToBytes();
        var options = new OfficeProvenanceOptions {
            MaxAssetBytes = pdf.Length + 1L,
            MaxManifestBytes = 64,
            MaxExpandedContainerBytes = 64
        };

        OfficeProvenanceReport report = PdfProvenance.Inspect(pdf, options);

        Assert.Single(report.Evidence);
        Assert.True(report.Evidence[0].IsStructurallyValid);
    }

    private static byte[] CreatePdfWithCandidateAndRetainedAttachment() => PdfDocument.Create()
        .Paragraph(paragraph => paragraph.Text("PDF provenance"))
        .AttachFile(new PdfEmbeddedFile(
            "keep.txt",
            Encoding.UTF8.GetBytes("keep"),
            "text/plain",
            PdfAssociatedFileRelationship.Supplement))
        .AttachFile(new PdfEmbeddedFile(
            "content-credential.c2pa",
            CreateManifestStore(),
            "application/c2pa",
            PdfAssociatedFileRelationship.C2paManifest))
        .ToBytes();

    private static byte[] DuplicateCandidateAroundRetainedAttachment(byte[] pdf, int copies) =>
        PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["EmbeddedFiles"]));
            PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items["Names"]));
            var pairs = new List<(PdfStringObj Name, PdfObject Reference)>();
            for (int index = 0; index + 1 < entries.Items.Count; index += 2) {
                pairs.Add((Assert.IsType<PdfStringObj>(entries.Items[index]), entries.Items[index + 1]));
            }
            (PdfStringObj Name, PdfObject Reference) candidate = pairs.Single(pair => pair.Name.Value.EndsWith(".c2pa", StringComparison.Ordinal));
            (PdfStringObj Name, PdfObject Reference) retained = pairs.Single(pair => pair.Name.Value == "keep.txt");
            entries.Items.Clear();
            for (int index = 0; index < copies; index++) {
                entries.Items.Add(new PdfStringObj(candidate.Name.Value + index, true));
                entries.Items.Add(candidate.Reference);
                if (index == 0) {
                    entries.Items.Add(retained.Name);
                    entries.Items.Add(retained.Reference);
                }
            }
            return security.InfoObjectNumber;
        });

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

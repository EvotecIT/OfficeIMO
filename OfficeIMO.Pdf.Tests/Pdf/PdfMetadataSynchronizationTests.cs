using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfMetadataSynchronizationTests {
    [Fact]
    public void GeneratedInfoDictionaryEncodesUnicodeWithoutPdfSyntaxInjection() {
        const string title = "\u0129 /Author \u0128Injected";

        byte[] pdf = PdfDocument.Create()
            .Meta(title: title)
            .Paragraph(paragraph => paragraph.Text("Metadata encoding"))
            .ToBytes();
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        string raw = PdfEncoding.Latin1GetString(pdf);

        Assert.Equal(title, info.Metadata.Title);
        Assert.Null(info.Metadata.Author);
        Assert.DoesNotContain(") /Author (Injected", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SynchronizeMetadata_UpdatesInfoAndXmpWhilePreservingCustomSchema() {
        var options = new PdfOptions { IncludeXmpMetadata = true }
            .SetElectronicInvoiceMetadata(new PdfElectronicInvoiceMetadata("ORDER", "invoice.xml", "1.0", "EN 16931"));
        byte[] source = PdfDocument.Create(options)
            .Meta(
                title: "Original title",
                author: "Original author",
                subject: "Preserved subject",
                keywords: "first; second")
            .Paragraph(paragraph => paragraph.Text("Full rewrite metadata synchronization"))
            .ToBytes();

        PdfMutationPlan plan = PdfMutationPlanner.Plan(
            source,
            PdfMutationOperation.SynchronizeMetadata,
            executionPreference: PdfMutationExecutionPreference.RequireFullRewrite);
        byte[] updated = PdfMetadataEditor.SynchronizeMetadata(
            source,
            title: "Synchronized title",
            author: string.Empty,
            keywords: "third, fourth");
        PdfDocumentInfo info = PdfInspector.Inspect(updated);

        Assert.Equal(PdfMutationExecutionMode.FullRewrite, plan.ExecutionMode);
        Assert.Contains(PdfMutationStructure.InfoDictionary, plan.AffectedStructures);
        Assert.Contains(PdfMutationStructure.XmpMetadata, plan.AffectedStructures);
        Assert.Contains(PdfMutationProof.MetadataReadback, plan.RequiredProofs);
        Assert.Equal("Synchronized title", info.Metadata.Title);
        Assert.Null(info.Metadata.Author);
        Assert.Equal("Preserved subject", info.Metadata.Subject);
        Assert.Equal("third, fourth", info.Metadata.Keywords);
        Assert.Equal("Synchronized title", info.XmpMetadata!.Title);
        Assert.Null(info.XmpMetadata.Creator);
        Assert.Equal("Preserved subject", info.XmpMetadata.Description);
        Assert.Equal(new[] { "third", "fourth" }, info.XmpMetadata.Subjects);
        Assert.Equal("ORDER", info.XmpMetadata.ElectronicInvoiceDocumentType);
        Assert.Equal("invoice.xml", info.XmpMetadata.ElectronicInvoiceDocumentFileName);
        Assert.Equal("EN 16931", info.XmpMetadata.ElectronicInvoiceConformanceLevel);
        Assert.Contains("Full rewrite metadata synchronization", PdfTextExtractor.ExtractAllText(updated), StringComparison.Ordinal);
        Assert.False(info.Security.HasIncrementalUpdates);
    }

    [Fact]
    public void FluentSynchronizeMetadata_CreatesXmpWhenMissing() {
        byte[] source = PdfDocument.Create()
            .Meta(title: "Info only", author: "Existing author")
            .Paragraph(paragraph => paragraph.Text("Create synchronized XMP"))
            .ToBytes();

        PdfDocument updated = PdfDocument.Open(source)
            .SynchronizeMetadata(title: "Created XMP", keywords: "one; two");
        PdfDocumentInfo info = updated.Inspect();

        Assert.Equal("Created XMP", info.Metadata.Title);
        Assert.Equal("Existing author", info.Metadata.Author);
        Assert.Equal("Created XMP", info.XmpMetadata!.Title);
        Assert.Equal("Existing author", info.XmpMetadata.Creator);
        Assert.Equal(new[] { "one", "two" }, info.XmpMetadata.Subjects);
    }

    [Fact]
    public void SynchronizeMetadata_CanPreserveMissingXmpByPolicy() {
        byte[] source = PdfDocument.Create()
            .Meta(title: "Info only")
            .Paragraph(paragraph => paragraph.Text("No XMP creation"))
            .ToBytes();

        byte[] updated = PdfMetadataEditor.SynchronizeMetadata(
            source,
            title: "Still Info only",
            createXmpMetadata: false);
        PdfDocumentInfo info = PdfInspector.Inspect(updated);

        Assert.Equal("Still Info only", info.Metadata.Title);
        Assert.Null(info.XmpMetadata);
    }

    [Fact]
    public void XmpSynchronizer_RejectsPacketsWithoutRdfInsteadOfDiscardingThem() {
        var metadata = new PdfMetadata { Title = "Unsafe replacement" };

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfXmpMetadataSynchronizer.Synchronize("<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" />", metadata));

        Assert.Contains("does not contain an RDF root", exception.Message, StringComparison.Ordinal);
    }
}

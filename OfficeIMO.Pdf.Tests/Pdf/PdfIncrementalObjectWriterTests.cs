using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfIncrementalObjectWriterTests {
    [Fact]
    public void Append_XrefStreamRevisionPreservesPrefixAndReadsUpdatedObject() {
        byte[] source = PdfDocument.Create()
            .Meta(title: "Original xref-stream title")
            .Paragraph(paragraph => paragraph.Text("Xref stream revision body"))
            .ToBytes();
        PdfDocumentSecurityInfo before = PdfSyntax.ReadDocumentSecurityInfo(source);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(source);
        int infoObjectNumber = objects.Keys.Max() + 1;
        byte[] infoObject = PdfObjectBytes.WrapIndirectObject(
            infoObjectNumber,
            PdfInfoDictionaryBuilder.Build(new PdfMetadata { Title = "Updated xref-stream title" }));

        byte[] updated = PdfIncrementalObjectWriter.Append(
            source,
            objects,
            before,
            trailerRaw,
            rawObjects: new[] { (infoObjectNumber, infoObject) },
            infoObjectNumberOverride: infoObjectNumber,
            format: PdfIncrementalXrefFormat.XrefStream);

        Assert.True(updated.AsSpan(0, source.Length).SequenceEqual(source));
        Assert.Equal("Updated xref-stream title", PdfInspector.Inspect(updated).Metadata.Title);
        PdfDocumentSecurityInfo after = PdfSyntax.ReadDocumentSecurityInfo(updated);
        Assert.True(after.HasXrefStreams);
        Assert.True(after.HasIncrementalUpdates);
        Assert.True(after.RevisionCount > before.RevisionCount);
        Assert.Equal(before.LastStartXrefOffset, after.Revisions[after.Revisions.Count - 1].PreviousXrefOffset);
    }

    [Fact]
    public void Append_ClassicRevisionRejectsDuplicateRawObjectNumbers() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Duplicate raw object source"))
            .ToBytes();
        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(source);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(source);
        int objectNumber = objects.Keys.Max() + 1;
        byte[] raw = PdfObjectBytes.WrapIndirectObject(objectNumber, PdfInfoDictionaryBuilder.Build(new PdfMetadata()));

        ArgumentException exception = Assert.Throws<ArgumentException>(() => PdfIncrementalObjectWriter.Append(
            source,
            objects,
            security,
            trailerRaw,
            rawObjects: new[] { (objectNumber, raw), (objectNumber, raw) },
            format: PdfIncrementalXrefFormat.ClassicTable));

        Assert.Contains("unique object numbers", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void UpdateMetadata_AppendsXrefStreamRevisionToXrefStreamInput() {
        byte[] source = PdfExternalDocumentCompatibilityTests.BuildXrefStreamPdfWithTrailingStaleDuplicatePage();
        PdfAppendOnlyMutationReport capability = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(source);
        PdfDocumentSecurityInfo sourceSecurity = PdfSyntax.ReadDocumentSecurityInfo(source);

        byte[] updated = PdfIncrementalUpdater.UpdateMetadata(source, title: "Xref stream public workflow");

        Assert.True(capability.CanAppendMetadata);
        Assert.False(sourceSecurity.BlocksOfficeIMOAppendOnlyMutation);
        Assert.True(updated.AsSpan(0, source.Length).SequenceEqual(source));
        Assert.Equal("Xref stream public workflow", PdfInspector.Inspect(updated).Metadata.Title);
        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(updated);
        Assert.True(security.HasXrefStreams);
        Assert.True(security.HasIncrementalUpdates);
        Assert.True(security.RevisionCount >= 2);
    }
}

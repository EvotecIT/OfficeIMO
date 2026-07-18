using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPipelineReportTests {
    [Fact]
    public void TryToBytes_ReportsExactGeneratedArtifact() {
        PdfDocument document = PdfDocument.Create()
            .Meta(title: "Pipeline")
            .Paragraph(paragraph => paragraph.Text("Generated artifact"));

        PdfBytesResult result = document.TryToBytes();

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        Assert.Collection(
            result.Pipeline.Steps,
            step => Assert.Equal(PdfPipelineStepKind.Create, step.Kind),
            step => {
                Assert.Equal(PdfPipelineStepKind.Output, step.Kind);
                Assert.Equal("ToBytes", step.Operation);
                Assert.True(step.Succeeded);
                Assert.NotNull(step.Duration);
            });
        Assert.Equal(result.ByteCount, result.Pipeline.Output?.ByteCount);
        Assert.Equal(1, result.Pipeline.Output?.PageCount);
        Assert.Equal(Sha256(result.Bytes), result.Pipeline.Output?.Sha256);
        Assert.Null(result.Pipeline.Input);
        Assert.True(result.Pipeline.TotalDuration >= TimeSpan.Zero);
    }

    [Fact]
    public void GeneratedStreamSave_ReportsWriterPageCountWithoutReadbackBuffering() {
        using var stream = new MemoryStream();

        PdfSaveResult result = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second page"))
            .Save(stream);

        Assert.True(result.Succeeded);
        Assert.Equal(2, result.Pipeline.Output?.PageCount);
        Assert.Equal(stream.Length, result.Pipeline.Output?.ByteCount);
        Assert.Equal(Sha256(stream.ToArray()), result.Pipeline.Output?.Sha256);
        Assert.Null(result.Pipeline.Input);
    }

    [Fact]
    public void GeneratedEncryptedBytes_ReportReadablePageCountWithWriterCredentials() {
        PdfBytesResult result = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Encrypted pipeline evidence"))
            .TryToBytes();

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        Assert.Equal(1, result.Pipeline.Output?.PageCount);
        Assert.Equal(
            1,
            PdfInspector.Inspect(result.Bytes, new PdfReadOptions { Password = "open" }).PageCount);
    }

    [Fact]
    public void OpenMutationAndSave_PreserveArtifactChainAndDecision() {
        byte[] source = PdfDocument.Create()
            .Meta(title: "Before")
            .Paragraph(paragraph => paragraph.Text("Pipeline mutation"))
            .ToBytes();

        PdfDocument updated = PdfDocument.Open(source).UpdateMetadata(title: "After");
        using var stream = new MemoryStream();

        PdfSaveResult result = updated.Save(stream);

        Assert.True(result.Succeeded);
        Assert.Equal(3, result.Pipeline.Steps.Count);
        PdfPipelineStep open = result.Pipeline.Steps[0];
        PdfPipelineStep mutation = result.Pipeline.Steps[1];
        PdfPipelineStep output = result.Pipeline.Steps[2];
        Assert.Equal(PdfPipelineStepKind.Open, open.Kind);
        Assert.Equal(open.Input?.Sha256, result.Pipeline.Input?.Sha256);
        Assert.Equal(PdfPipelineStepKind.Mutation, mutation.Kind);
        Assert.Equal(PdfMutationOperation.UpdateMetadata, mutation.MutationOperation);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, mutation.ExecutionMode);
        Assert.Equal(open.Output?.Sha256, mutation.Input?.Sha256);
        Assert.Equal(mutation.Output?.Sha256, output.Input?.Sha256);
        Assert.Equal(Sha256(stream.ToArray()), output.Output?.Sha256);
        Assert.Equal(1, output.Output?.PageCount);
        Assert.Equal("After", PdfInspector.Inspect(stream.ToArray()).Metadata.Title);
    }

    [Fact]
    public void AppendRevision_ReportsAppendOnlyMutation() {
        byte[] source = PdfDocument.Create()
            .Meta(title: "Before")
            .Paragraph(paragraph => paragraph.Text("Append-only pipeline"))
            .ToBytes();

        PdfDocument updated = PdfDocument.Open(source)
            .AppendMetadataRevision(title: "After");

        PdfPipelineStep mutation = Assert.Single(
            updated.Pipeline.Steps,
            step => step.Kind == PdfPipelineStepKind.Mutation);
        Assert.Equal(PdfMutationOperation.UpdateMetadata, mutation.MutationOperation);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, mutation.ExecutionMode);
        Assert.Equal(source.LongLength, mutation.Input?.ByteCount);
        Assert.True(mutation.Output?.ByteCount > source.LongLength);
    }

    [Fact]
    public void TrySaveFailure_ReportsFailedOutputStage() {
        using var stream = new MemoryStream(new byte[16], writable: false);

        PdfSaveResult result = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Failure evidence"))
            .TrySave(stream);

        Assert.False(result.Succeeded);
        Assert.True(result.Pipeline.HasFailures);
        PdfPipelineStep output = Assert.Single(
            result.Pipeline.Steps,
            step => step.Kind == PdfPipelineStepKind.Output);
        Assert.False(output.Succeeded);
        Assert.NotEmpty(output.Diagnostics);
    }

    [Fact]
    public void AdapterFailureFactory_ReportsFailedOutputPipeline() {
        var exception = new InvalidOperationException("Adapter conversion failed.");

        PdfSaveResult result = PdfSaveResult.FromFailure("failed.pdf", exception);

        Assert.False(result.Succeeded);
        Assert.True(result.Pipeline.HasFailures);
        PdfPipelineStep output = Assert.Single(result.Pipeline.Steps);
        Assert.Equal(PdfPipelineStepKind.Output, output.Kind);
        Assert.Equal("Save", output.Operation);
        Assert.False(output.Succeeded);
        Assert.Contains(output.Diagnostics, diagnostic => diagnostic.Contains("Adapter conversion failed.", StringComparison.Ordinal));
    }

    [Fact]
    public void GeneratedAppendMutation_ReportsTheExactConsumedInputArtifact() {
        PdfDocument updated = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Generated mutation input"))
            .AppendMetadataRevision(title: "After");

        PdfPipelineStep mutation = Assert.Single(
            updated.Pipeline.Steps,
            step => step.Kind == PdfPipelineStepKind.Mutation);
        byte[] output = updated.ToBytes();
        Assert.NotNull(mutation.Input);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, mutation.ExecutionMode);
        Assert.True(output.LongLength > mutation.Input!.ByteCount);

        byte[] consumedInput = output
            .Take(checked((int)mutation.Input.ByteCount))
            .ToArray();
        Assert.Equal(Sha256(consumedInput), mutation.Input.Sha256);
    }

    [Fact]
    public void PageImportAndCropMutations_PreserveOperationClassification() {
        byte[] importedPage = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Imported"))
            .ToBytes();
        PdfDocument appended = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Target"))
            .Pages.Append(importedPage);
        PdfDocument cropped = appended.Pages.CropAndTranslate(0, 0, 300, 300, 1);

        PdfPipelineStep append = Assert.Single(
            appended.Pipeline.Steps,
            step => step.Kind == PdfPipelineStepKind.Mutation);
        PdfPipelineStep crop = cropped.Pipeline.Steps.Last(
            step => step.Kind == PdfPipelineStepKind.Mutation);
        Assert.Equal("Append", append.Operation);
        Assert.Equal(PdfMutationOperation.MergeDocuments, append.MutationOperation);
        Assert.Equal("CropAndTranslate", crop.Operation);
        Assert.Equal(PdfMutationOperation.ModifyPageTree, crop.MutationOperation);
    }

    [Fact]
    public void PageSplitOutputs_PreserveSourceLineageAndExtractionClassification() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Third"))
            .ToBytes();
        PdfDocument opened = PdfDocument.Open(source);

        IReadOnlyList<PdfDocument> pages = opened.Pages.Split();
        PdfDocument range = Assert.Single(opened.Pages.Split(new[] { PdfPageRange.From(1, 2) }));

        Assert.All(pages, part => AssertSplitLineage(part, source, expectedPageCount: 1));
        AssertSplitLineage(range, source, expectedPageCount: 2);
    }

    [Fact]
    public void GeneratedPageSplitOutputs_ReuseOneCapturedInputArtifact() {
        PdfDocument generated = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Third"));

        IReadOnlyList<PdfDocument> pages = generated.Pages.Split();

        PdfArtifactSnapshot input = Assert.IsType<PdfArtifactSnapshot>(pages[0].Pipeline.Steps[1].Input);
        Assert.All(
            pages,
            page => Assert.Same(input, page.Pipeline.Steps[1].Input));
    }

    private static void AssertSplitLineage(PdfDocument document, byte[] source, int expectedPageCount) {
        Assert.Collection(
            document.Pipeline.Steps,
            open => {
                Assert.Equal(PdfPipelineStepKind.Open, open.Kind);
                Assert.Equal(Sha256(source), open.Input?.Sha256);
            },
            mutation => {
                Assert.Equal(PdfPipelineStepKind.Mutation, mutation.Kind);
                Assert.Equal("Split", mutation.Operation);
                Assert.Equal(PdfMutationOperation.ExtractPages, mutation.MutationOperation);
                Assert.Equal(Sha256(source), mutation.Input?.Sha256);
                Assert.Equal(expectedPageCount, mutation.Output?.PageCount);
            });
    }

    private static string Sha256(byte[] bytes) => PdfArtifactFingerprint.ComputeSha256(bytes);
}

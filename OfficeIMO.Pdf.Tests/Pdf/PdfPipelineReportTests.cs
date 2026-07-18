using System.Security.Cryptography;
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

    private static string Sha256(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}

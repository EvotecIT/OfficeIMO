using OfficeIMO.OpenDocument;

namespace OfficeIMO.Workflows.Tests;

public sealed class ConversionPublicationEvidenceTests {
    [Theory]
    [InlineData("pub", "conflict")]
    [InlineData("pub", "guard")]
    [InlineData("fodg", "conflict")]
    [InlineData("fodg", "guard")]
    [InlineData("pub", "cancel")]
    public async Task CompletedConversionEvidenceSurvivesPublicationFailure(string extension, string failure) {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-conversion-evidence-" + Guid.NewGuid().ToString("N"))).FullName;
        try {
            string input = Path.Combine(root, "source." + extension), output = Path.Combine(root, "result.pdf");
            if (extension == "pub") File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Publisher", "Sample.pub"), input);
            else {
                OdgDocument source = OdgDocument.Create();
                source.AddPage("First"); source.AddPage("Second"); source.SaveFlatXml(input);
            }
            byte[] sentinel = [1, 2, 3]; File.WriteAllBytes(output, sentinel);
            using var cancellation = new CancellationTokenSource();
            OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new() {
                Operation = OfficeWorkflowOperation.Convert, InputPath = input, OutputPath = output,
                ConversionRouteId = extension == "pub" ? "publisher-pdf" : "odg-pdf",
                ConflictPolicy = failure == "conflict" ? OfficeWorkflowConflictPolicy.Fail : OfficeWorkflowConflictPolicy.Replace,
                PublicationGuard = failure == "guard" ? new DenyGuard() : null
            }, progress: failure == "cancel" ? new CancelAtPublication(cancellation) : null, cancellationToken: cancellation.Token);
            Assert.False(result.Succeeded);
            Assert.Equal(failure == "cancel" ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed, result.Status);
            Assert.True(result.FailureKind == (failure == "cancel" ? OfficeWorkflowFailureKind.None : OfficeWorkflowFailureKind.OutputFailed), result.Summary);
            Assert.Equal(sentinel, File.ReadAllBytes(output)); Assert.Empty(Directory.GetFiles(root, ".*.tmp"));
            OfficeWorkflowConversionEvidence evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
            Assert.Equal(extension.ToUpperInvariant(), evidence.Facts["sourceFormat"]); Assert.Equal("2", evidence.Facts["sourcePages"]);
            if (extension == "pub") { Assert.True(evidence.HasLoss); Assert.NotEmpty(evidence.FidelityDiagnostics); }
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OutputReopened");
        } finally { Directory.Delete(root, true); }
    }

    private sealed class DenyGuard : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) => new(false);
    }

    private sealed class CancelAtPublication(CancellationTokenSource cancellation) : IProgress<OfficeWorkflowProgress> {
        public void Report(OfficeWorkflowProgress value) { if (value.Stage == "publish") cancellation.Cancel(); }
    }
}

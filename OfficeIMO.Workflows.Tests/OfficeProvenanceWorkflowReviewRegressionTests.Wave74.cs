namespace OfficeIMO.Workflows.Tests;

public sealed partial class OfficeProvenanceWorkflowTests {
    [Fact]
    public async Task SuccessfulRemovalReportsCompletionAfterPublicationFinalizes() {
        using var scope = new TempScope();
        string input = scope.Write("page.html", HtmlWithExternalManifest("original"));
        string output = Path.Combine(scope.Path, "cleaned.html");
        var stages = new List<string>();

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = input,
                OutputPath = output
            },
            new RecordingProgress(stages));

        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal("complete", stages[^1]);
        Assert.True(stages.IndexOf("finalize") < stages.IndexOf("complete"));
        Assert.True(File.Exists(output));
    }

    private sealed class CancellingFinalizationProgress : IProgress<OfficeWorkflowProgress> {
        private readonly string _requestId;
        private readonly CancellationTokenSource _cancellation;

        internal CancellingFinalizationProgress(string requestId, CancellationTokenSource cancellation) {
            _requestId = requestId;
            _cancellation = cancellation;
        }

        internal List<string> Stages { get; } = new();

        public void Report(OfficeWorkflowProgress value) {
            Stages.Add(value.Stage);
            if (value.RequestId == _requestId && value.Stage == "finalize") _cancellation.Cancel();
        }
    }

    private sealed class RecordingProgress(List<string> stages) : IProgress<OfficeWorkflowProgress> {
        public void Report(OfficeWorkflowProgress value) => stages.Add(value.Stage);
    }
}

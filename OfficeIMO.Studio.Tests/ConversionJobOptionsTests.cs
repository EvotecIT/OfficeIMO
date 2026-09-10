using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class ConversionJobOptionsTests {
    [Fact]
    public async Task QueuedProfilesAndRunningSettingsAreIndependentOfLaterChanges() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-job-options-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string word = Path.Combine(root, "input.docx"), pdf = Path.Combine(root, "input.pdf");
            File.WriteAllBytes(word, [1]); File.WriteAllBytes(pdf, [1]);
            IReadOnlyList<string> files = [word];
            var runner = new CapturingRunner();
            using var model = new ConversionWorkbenchViewModel(_ => Task.FromResult(files), _ => Task.FromResult<string?>(root), runner);
            model.SelectedRoute = model.Routes.Single(route => route.Route.Id == "docx-pdf");
            model.SelectedProfile = model.AvailableProfiles.Single(profile => profile.Value == OfficeWorkflowOutputProfile.PrintReady);
            await model.AddFilesCommand.ExecuteAsync(null);
            files = [pdf];
            model.SelectedRoute = model.Routes.Single(route => route.Route.Id == "pdf-docx");
            Assert.Single(model.AvailableProfiles);
            await model.AddFilesCommand.ExecuteAsync(null);
            var pdfJob = model.Jobs[1];
            pdfJob.PageRanges = "2"; pdfJob.WordMode = PdfWordImportMode.VisualPages; pdfJob.RasterDpi = 72;
            Task run = model.RunQueueCommand.ExecuteAsync(null);
            await runner.Started.Task.WaitAsync(TimeSpan.FromSeconds(10));
            pdfJob.PageRanges = "3"; pdfJob.WordMode = PdfWordImportMode.EditableContent;
            model.SelectedProfile = model.Profiles.Single(profile => profile.Value == OfficeWorkflowOutputProfile.TextOnly);
            Assert.Equal(OfficeWorkflowOutputProfile.PrintReady, runner.Requests[0].OutputProfile);
            Assert.Equal(OfficeWorkflowOutputProfile.Faithful, runner.Requests[1].OutputProfile);
            Assert.Equal("2", runner.Requests[1].ConversionOptions!.PageRanges);
            Assert.Equal(PdfWordImportMode.VisualPages, runner.Requests[1].ConversionOptions!.WordMode);
            Assert.Equal(72, runner.Requests[1].ConversionOptions!.RasterDpi);
            runner.Completed.SetResult([]);
            await run;
        } finally { Directory.Delete(root, true); }
    }

    private sealed class CapturingRunner : IOfficeWorkflowRunner {
        internal OfficeWorkflowRequest[] Requests = [];
        internal TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource<IReadOnlyList<OfficeWorkflowResult>> Completed { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public Task<OfficeWorkflowResult> RunAsync(OfficeWorkflowRequest request, IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default) => throw new NotSupportedException();
        public Task<IReadOnlyList<OfficeWorkflowResult>> RunBatchAsync(IEnumerable<OfficeWorkflowRequest> requests, IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default) {
            Requests = requests.ToArray(); Started.SetResult(); return Completed.Task;
        }
    }
}

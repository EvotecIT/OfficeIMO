using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioConversionLossNoticeTests {
    [Fact]
    public async Task Lossy_outputs_are_named_in_one_notice_that_leads_to_their_details() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-loss-notice-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            // Publisher pages are drawn with documented approximations, so the conversion completes and reports loss.
            string source = Path.Combine(root, "newsletter.pub");
            File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Publisher", "Sample.pub"), source);
            IReadOnlyList<string> files = [source];
            using var model = new ConversionWorkbenchViewModel(_ => Task.FromResult(files), _ => Task.FromResult<string?>(root), new OfficeWorkflowRunner());
            model.SelectedRoute = model.Routes.Single(route => route.Route.Id == "publisher-pdf");
            await model.AddFilesCommand.ExecuteAsync(null);
            Assert.False(model.HasLossNotice);

            await model.RunQueueCommand.ExecuteAsync(null);

            ConversionJobViewModel job = Assert.Single(model.Jobs);
            Assert.True(job.State == ConversionJobState.Completed, job.Summary);
            Assert.True(job.HasLoss);
            Assert.True(model.HasLossNotice);
            Assert.Contains("newsletter.pub", model.LossNotice, StringComparison.Ordinal);

            model.SelectedJob = null;
            model.ReviewLossCommand.Execute(null);
            Assert.Same(job, model.SelectedJob);

            // Removing the job takes it out of the notice.
            model.RemoveSelectedCommand.Execute(null);
            Assert.False(model.HasLossNotice);
        } finally { Directory.Delete(root, true); }
    }
}

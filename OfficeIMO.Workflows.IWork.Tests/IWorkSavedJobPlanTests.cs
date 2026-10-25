using System.IO.Compression;
using OfficeIMO.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed class IWorkSavedJobPlanTests {
    [Theory]
    [InlineData("pages", "docx")]
    [InlineData("numbers", "xlsx")]
    [InlineData("key", "pptx")]
    public async Task SavedJobPlanAdmitsOnlyRegisteredDirectoryPackagesAndPreservesCheckpointRestriction(string kind, string target) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-saved-package-plan-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "source." + kind);
        string output = Path.Combine(root, "output");
        string state = Path.Combine(root, "state");
        Directory.CreateDirectory(input);
        try {
            ZipFile.ExtractToDirectory(Path.Combine(AppContext.BaseDirectory, "Corpus", kind == "key" ? "tabledeck.key" : "simple." + kind), input);
            var job = new OfficeConversionJob { InputPaths = [input], OutputDirectory = output, TargetExtension = "." + target };
            var runner = IWorkWorkflow.CreateRunner(conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
            OfficeConversionJobPlan plan = await OfficeConversionJobRunner.PlanAsync(job, root, runner);
            Assert.True(plan.CanRun, string.Join("; ", plan.Diagnostics.Concat(plan.Items.Select(item => item.Message))));
            Assert.Equal(1, plan.Selected); Assert.Equal(0, plan.Failed);
            Assert.Contains("Directory package admitted", Assert.Single(plan.Items).Message);
            Assert.False(Directory.Exists(output)); Assert.False(Directory.Exists(state));
            string routeId = Assert.IsType<string>(Assert.Single(plan.Items).RouteId);
            var fileOnlyRunner = new OfficeWorkflowRunner(null, null, [new OfficeWorkflowConversionRegistration(routeId,
                (_, _, _, _) => throw new InvalidOperationException("Preflight must not execute the converter."))]);
            OfficeConversionJobPlan fileOnlyPlan = await OfficeConversionJobRunner.PlanAsync(job, root, fileOnlyRunner);
            Assert.False(fileOnlyPlan.CanRun);
            Assert.Contains(fileOnlyPlan.Items, item => item.Disposition == "Blocked" && item.Message.Contains("does not admit directory packages"));
            job.CheckpointDirectory = state;
            OfficeConversionJobPlan checkpointPlan = await OfficeConversionJobRunner.PlanAsync(job, root, runner);
            Assert.False(checkpointPlan.CanRun);
            Assert.Contains(checkpointPlan.Items, item => item.Disposition == "Blocked" && item.Message.Contains("without checkpoints"));
            Assert.False(Directory.Exists(output)); Assert.False(Directory.Exists(state));
            job.CheckpointDirectory = null;
            OfficeConversionBatchResult result = await OfficeConversionJobRunner.RunAsync(job, root, runner);
            Assert.Equal(1, result.Completed); Assert.Equal(0, result.Failed); Assert.Null(result.FailureSummary);
            Assert.True(File.Exists(Path.Combine(output, "source." + kind + "." + target)));
        } finally { if (Directory.Exists(root)) Directory.Delete(root, recursive: true); }
    }
}

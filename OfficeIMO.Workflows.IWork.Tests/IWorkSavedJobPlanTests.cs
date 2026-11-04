using System.IO.Compression;
using OfficeIMO.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed class IWorkSavedJobPlanTests {
    // The executor boundary proves planning rejects a file-only implementation without executing it.
    private sealed class FileOnlyExecutor : IOfficeConversionExecutor {
        public IReadOnlyList<OfficeConversionExecutionRoute> ConversionRoutes => new OfficeConversionExecutor().ConversionRoutes;
        public bool SupportsDirectoryPackageInput(string routeId) => false;
        public Task<OfficeConversionExecutionResult> RunAsync(OfficeConversionExecutionRequest request, CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException("Preflight must not execute the converter.");
    }

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
            var job = new OfficeBatchOptions { InputPaths = [input], OutputDirectory = output, TargetExtension = "." + target };
            IWorkWorkflow.RegisterConversions();
            job = OfficeBatchOptions.Parse(job.ToJson(root), root);
            OfficeBatchPlan plan = await OfficeConverter.PlanAsync(job);
            Assert.True(plan.CanRun, string.Join("; ", plan.Diagnostics.Concat(plan.Items.Select(item => item.Message))));
            Assert.Equal(1, plan.Selected); Assert.Equal(0, plan.Failed);
            Assert.Contains("Directory package admitted", Assert.Single(plan.Items).Message);
            Assert.False(Directory.Exists(output)); Assert.False(Directory.Exists(state));
            string routeId = Assert.IsType<string>(Assert.Single(plan.Items).RouteId);
            OfficeBatchPlan fileOnlyPlan = await OfficeBatchExecutor.PlanAsync(new FileOnlyExecutor(), job, null, default);
            Assert.False(fileOnlyPlan.CanRun);
            Assert.Contains(fileOnlyPlan.Items, item => item.Disposition == "Blocked" && item.Message.Contains("does not admit directory packages"));
            job.ResumeFrom = Path.Combine(state, "plan.json");
            OfficeBatchPlan checkpointPlan = await OfficeConverter.PlanAsync(job);
            Assert.False(checkpointPlan.CanRun);
            Assert.Contains(checkpointPlan.Items, item => item.Disposition == "Blocked" && item.Message.Contains("without checkpoints"));
            Assert.False(Directory.Exists(output)); Assert.False(Directory.Exists(state));
            job.ResumeFrom = null;
            job.ConversionOptions.RegisteredSettings = new IWorkWorkflowSettings { ConversionOptions = new IWorkConversionOptions { AllowPartialEditableReconstruction = true } };
            OfficeBatchSummary result = (await OfficeConverter.ConvertAsync(job)).Summary;
            Assert.Equal(1, result.Completed); Assert.Equal(0, result.Failed); Assert.Null(result.FailureSummary);
            Assert.True(File.Exists(Path.Combine(output, "source." + target)));
        } finally { if (Directory.Exists(root)) Directory.Delete(root, recursive: true); }
    }
}

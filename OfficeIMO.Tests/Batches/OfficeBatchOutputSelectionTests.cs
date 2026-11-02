using OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class OfficeBatchOutputSelectionTests {
    [Theory]
    [InlineData(true, -1, false)]
    [InlineData(true, -1, true)]
    [InlineData(false, 1, false)]
    [InlineData(false, 1, true)]
    [InlineData(false, 0, false)]
    [InlineData(false, -1, false)]
    [InlineData(false, 2, false)]
    public async Task PlansAndRunsShareOutputSelectionWithoutClaimingConversionOrReuse(bool noClobber, int outputAge, bool checkpoint) {
        string root = Directory.CreateTempSubdirectory("officeimo-output-selection-").FullName;
        try {
            string source = Path.Combine(root, "review.md");
            string outputDirectory = Directory.CreateDirectory(Path.Combine(root, "pdf")).FullName;
            string destination = Path.Combine(outputDirectory, "review.pdf");
            File.WriteAllText(source, "# Output selection\n\nNew conversion evidence.");
            DateTime sourceTime = new(2025, 1, 2, 0, 0, 0, DateTimeKind.Utc);
            File.SetLastWriteTimeUtc(source, sourceTime);
            byte[] original = [1, 2, 3, 4];
            bool exists = outputAge != 2;
            if (exists) {
                File.WriteAllBytes(destination, original);
                File.SetLastWriteTimeUtc(destination, sourceTime.AddDays(outputAge));
            }
            var job = new OfficeBatchOptions {
                InputPaths = [source], OutputDirectory = outputDirectory,
                ConflictPolicy = checkpoint || noClobber ? OfficePublicationPolicy.Fail : OfficePublicationPolicy.Replace,
                SkipExistingOutputs = noClobber, SkipUpToDate = true,
                ResumeFrom = checkpoint ? Path.Combine(root, "state", "checkpoint.json") : null
            };
            // Exercise the portable contract used by saved jobs before either preview or execution.
            job = OfficeBatchOptions.Parse(job.ToJson(root), root);
            Assert.Equal(noClobber, job.SkipExistingOutputs);
            Assert.True(job.SkipUpToDate);
            bool shouldSkip = exists && (noClobber || outputAge == 1);
            OfficeBatchPlan plan = await OfficeConverter.PlanAsync(job);
            Assert.True(plan.CanRun);
            Assert.Equal(1, plan.Selected);
            Assert.Equal(shouldSkip ? 1 : 0, plan.Skipped);
            Assert.Equal(0, plan.Failed);
            Assert.Equal(shouldSkip ? "Skipped" : "Selected", Assert.Single(plan.Items).Disposition);

            OfficeBatchReport report = await OfficeConverter.ConvertAsync(job);
            Assert.True(report.Succeeded);
            Assert.Equal(1, report.Summary.Selected);
            Assert.Equal(shouldSkip ? 1 : 0, report.Summary.Skipped);
            Assert.Equal(shouldSkip ? 0 : 1, report.Summary.Completed);
            Assert.Equal(0, report.Summary.Reused);
            Assert.Equal(0, report.Summary.Failed);
            OfficeBatchItemReport item = Assert.Single(report.Items);
            if (shouldSkip) {
                Assert.Equal(noClobber ? "OutputExists" : "UpToDate", item.SkippedReason);
                Assert.Null(item.Status);
                Assert.Null(item.OutputPath);
                Assert.Null(item.Report);
                Assert.Equal(original, File.ReadAllBytes(destination));
                Assert.Equal(sourceTime.AddDays(outputAge), File.GetLastWriteTimeUtc(destination));
                Assert.Equal(item.SkippedReason, Assert.Single(OfficeBatchReport.Parse(report.ToJson()).Items).SkippedReason);
            } else {
                using var pdf = PdfDocument.Load(destination);
                Assert.Contains("New conversion evidence.", pdf.Read().Text);
                Assert.Equal(OfficeBatchItemStatus.Completed, item.Status);
                Assert.Null(item.SkippedReason);
            }
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task SkippingExistingOutputsDoesNotTurnUnsupportedSelectionsIntoSuccessfulWork() {
        string root = Directory.CreateTempSubdirectory("officeimo-unselected-output-").FullName;
        try {
            string source = Path.Combine(root, "unsupported.bin");
            File.WriteAllText(source, "Unsupported source");
            var report = await OfficeConverter.ConvertAsync(new OfficeBatchOptions {
                InputPaths = [source], OutputDirectory = Path.Combine(root, "pdf"), SkipExistingOutputs = true
            });
            Assert.False(report.Succeeded);
            Assert.Equal(0, report.Summary.Selected);
            Assert.Equal(1, report.Summary.Skipped);
            Assert.Null(Assert.Single(report.Items).SkippedReason);
        } finally { Directory.Delete(root, true); }
    }
}

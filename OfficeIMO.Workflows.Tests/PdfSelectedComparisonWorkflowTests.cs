using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfSelectedComparisonWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BothCapturedComparisonSourcesAreReverifiedBeforeReportPublication(bool changeActual) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-selected-publication-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string expected = Path.Combine(root, "expected.pdf"), actual = Path.Combine(root, "actual.pdf"), output = Path.Combine(root, "report.html");
            PdfDocument.Create(document => document.Page(page => page.Size(150, 180))).Save(expected);
            File.Copy(expected, actual); File.WriteAllText(output, "existing report");
            OfficeWorkflowStreamInput Source(string path) => new(Path.GetFileName(path), _ => Task.FromResult<Stream>(File.OpenRead(path)));
            var result = await new OfficeWorkflowRunner().RunAsync(new() {
                Operation = OfficeWorkflowOperation.Compare, InputPath = expected, ComparisonPath = actual,
                InputStream = Source(expected), ComparisonStream = Source(actual), OutputPath = output,
                ComparisonExpectedPages = PdfPageSelector.Parse("1"), ComparisonActualPages = PdfPageSelector.Parse("1"),
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                PublicationGuard = new Guard(() => PdfDocument.Create(document => document.Page(page => page.Size(170, 200))).Save(changeActual ? actual : expected))
            });
            Assert.False(result.Succeeded); Assert.Equal("existing report", File.ReadAllText(output));
        } finally { Directory.Delete(root, true); }
    }

    private sealed class Guard(Action beforePublication) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) {
            beforePublication(); return ValueTask.FromResult(true);
        }
    }

    [Fact]
    public async Task CapturedScopesProduceStandaloneReportAndRejectStaleSourcesAndOutputAliases() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-selected-comparison-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string expected = Path.Combine(root, "expected.pdf"), actual = Path.Combine(root, "actual.pdf"), output = Path.Combine(root, "report.html");
            PdfDocument.Create(document => { for (int i = 0; i < 110; i++) document.Page(page => page.Size(150, 180)); }).Save(expected);
            File.Copy(expected, actual);
            byte[] before = File.ReadAllBytes(expected);
            var request = OfficeWorkflow.Compare(expected, actual).ComparePages(PdfPageSelector.Parse("110,2"), PdfPageSelector.Parse("2,110,1"))
                .To(output).Build();
            var result = await new OfficeWorkflowRunner().RunAsync(request);
            Assert.True(result.Succeeded, result.Summary);
            Assert.Equal("110", result.HealthReport!.Metrics["expectedTotalPages"]);
            Assert.Equal("2", result.HealthReport.Metrics["expectedSelectedPageCount"]);
            Assert.Equal("3", result.HealthReport.Metrics["actualSelectedPageCount"]);
            Assert.Equal("110,2", result.HealthReport.Metrics["expectedSelectedPages"]);
            Assert.Equal("1", result.HealthReport.Metrics["unmatchedActualPages"]);
            Assert.Contains("Expected 110 / actual 2", File.ReadAllText(output));
            Assert.Equal(before, File.ReadAllBytes(expected)); Assert.Equal(before, File.ReadAllBytes(actual));
            request.OutputPath = expected;
            Assert.False((await new OfficeWorkflowRunner().RunAsync(request)).Succeeded);
            string staleOutput = Path.Combine(root, "stale.html"); request.OutputPath = staleOutput;
            request.InputStream = new OfficeWorkflowStreamInput("expected.pdf", _ => Task.FromResult<Stream>(File.OpenRead(expected)),
                expectedSha256: new string('0', 64));
            Assert.False((await new OfficeWorkflowRunner().RunAsync(request)).Succeeded);
            Assert.False(File.Exists(staleOutput)); Assert.Equal(before, File.ReadAllBytes(expected));
        } finally { Directory.Delete(root, true); }
    }
}

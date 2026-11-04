using OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterBatchTests {
    [Fact]
    public async Task ManifestInsideOutputReusesOnlyVerifiedArtifacts() {
        using var files = new BatchFiles();
        string source = files.Write("source.txt", "Durable conversion text");
        OfficeBatchOptions settings = files.Settings(source);
        OfficeBatchPlan plan = await OfficeConverter.PlanAsync(settings);
        Assert.True(plan.CanRun);
        Assert.False(Directory.Exists(settings.OutputDirectory));
        OfficeBatchReport first = await OfficeConverter.ConvertAsync(settings);
        Assert.True(first.Succeeded, string.Join("; ", first.Items.Select(item => item.Summary)));
        Assert.Equal(1, first.Summary.Completed);
        Assert.Equal(0, first.Summary.Reused);
        Assert.True(File.Exists(settings.ResumeFrom));
        OfficeBatchItemReport item = Assert.Single(first.Items);
        string output = item.OutputPath!;
        Assert.Equal(Path.Combine(settings.OutputDirectory, "source.pdf"), output);
        Assert.Equal(output, Assert.IsType<OfficeConversionReport>(item.Report).OutputPath);
        OfficeBatchReport saved = OfficeBatchReport.Parse(first.ToJson());
        Assert.Equal(output, Assert.IsType<OfficeConversionReport>(Assert.Single(saved.Items).Report).OutputPath);
        using (var pdf = PdfDocument.Load(output)) Assert.Contains("Durable conversion text", pdf.ExtractText());
        byte[] original = File.ReadAllBytes(output);
        OfficeBatchReport resumed = await OfficeConverter.ConvertAsync(settings);
        Assert.True(resumed.Succeeded);
        Assert.Equal(1, resumed.Summary.Reused);
        Assert.Equal(original, File.ReadAllBytes(output));
    }

    [Fact]
    public async Task ChangedSourceAndOutputCannotReplaceACompletionReceipt() {
        using var files = new BatchFiles();
        string source = files.Write("source.txt", "Original input");
        OfficeBatchOptions settings = files.Settings(source);
        OfficeBatchReport first = await OfficeConverter.ConvertAsync(settings);
        Assert.True(first.Succeeded);
        string output = Assert.Single(first.Items).OutputPath!;
        string receipt = Assert.Single(Directory.GetFiles(settings.ResumeFrom + ".state/items", "*.json", SearchOption.AllDirectories));
        byte[] originalReceipt = File.ReadAllBytes(receipt), originalOutput = File.ReadAllBytes(output);
        File.WriteAllText(source, "Changed input");
        OfficeBatchReport changed = await OfficeConverter.ConvertAsync(settings);
        Assert.False(changed.Succeeded);
        Assert.Equal(1, changed.Summary.Failed);
        Assert.Equal(originalReceipt, File.ReadAllBytes(receipt));
        Assert.Equal(originalOutput, File.ReadAllBytes(output));
        File.WriteAllText(source, "Original input");
        File.WriteAllText(output, "Changed output");
        OfficeBatchReport damaged = await OfficeConverter.ConvertAsync(settings);
        Assert.False(damaged.Succeeded);
        Assert.Equal(originalReceipt, File.ReadAllBytes(receipt));
        Assert.Equal("Changed output", File.ReadAllText(output));
    }

    [Fact]
    public async Task DurableIntentRecoversAfterPublicationIsDenied() {
        using var files = new BatchFiles();
        string source = files.Write("source.txt", "Recoverable output");
        OfficeBatchOptions settings = files.Settings(source);
        string output = Path.Combine(settings.OutputDirectory, "source.pdf");
        var guard = new DenyFinalPublication(output);
        OfficeBatchReport denied = await OfficeConverter.ConvertAsync(settings with { PublicationGuard = guard });
        Assert.False(denied.Succeeded);
        Assert.False(File.Exists(output));
        Assert.NotEmpty(Directory.GetFiles(settings.OutputDirectory, ".*.conversion.pdf"));
        OfficeBatchReport recovered = await OfficeConverter.ConvertAsync(settings);
        Assert.True(recovered.Succeeded, string.Join("; ", recovered.Items.Select(item => item.Summary)));
        Assert.Equal(1, recovered.Summary.Reused);
        Assert.True(File.Exists(output));
        Assert.Empty(Directory.GetFiles(settings.OutputDirectory, ".*.conversion.pdf"));
    }

    [Fact]
    public async Task MixedSelectionReportsSkippedAndFailedItemsWithoutLosingCounts() {
        using var files = new BatchFiles();
        string valid = files.Write("source.txt", "Valid text");
        string ignored = files.Write("ignored.unknown", "Unselected");
        string invalid = files.Write("invalid.docx", "Not an Office package");
        OfficeBatchReport report = await OfficeConverter.ConvertAsync([valid, ignored, invalid], Path.Combine(files.Root, "output"),
            OfficeDocumentFormat.Pdf, new OfficeBatchOptions { MaxParallelism = 2 });
        Assert.False(report.Succeeded);
        Assert.Equal(2, report.Summary.Selected);
        Assert.Equal(1, report.Summary.Completed);
        Assert.Equal(1, report.Summary.Failed);
        Assert.Equal(1, report.Summary.Skipped);
        Assert.Equal(3, report.Items.Count);
        string json = OfficeBatchReportSerializer.Serialize(report);
        Assert.Equal(report.Summary, OfficeBatchReportSerializer.Deserialize(json).Summary);
    }

    private sealed class DenyFinalPublication(string output) : IOfficePublicationGuard {
        private int _visits;
        public Task<bool> CanPublishAsync(string absoluteDestination, bool isDirectory, CancellationToken cancellationToken) =>
            Task.FromResult(absoluteDestination != output || Interlocked.Increment(ref _visits) == 1);
    }

    private sealed class BatchFiles : IDisposable {
        internal string Root { get; } = Directory.CreateTempSubdirectory("officeimo-batch-test-").FullName;
        internal string Write(string name, string text) {
            string directory = Path.Combine(Root, "input"); Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, name); File.WriteAllText(path, text); return path;
        }
        internal OfficeBatchOptions Settings(string source) => new() { InputPaths = [source], OutputDirectory = Path.Combine(Root, "output"),
            ResumeFrom = Path.Combine(Root, "output", ".officeimo-batch.json") };
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }
}

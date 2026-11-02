using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterBatchNamingTests {
    [Theory]
    [InlineData("Sales.docx", "Sales.pdf")]
    [InlineData("Sales.v2.DOCX", "Sales.v2.pdf")]
    public async Task WordToPdfPlanAndConversionReplaceOnlyTheFinalExtension(string name, string expectedName) {
        using var files = new NamingFiles();
        string source = files.PathInInput(name);
        using (var word = WordDocument.Create()) { word.AddParagraph("Sales document content"); word.Save(source); }
        byte[] sourceBytes = File.ReadAllBytes(source);
        OfficeBatchOptions settings = files.Settings(source);
        OfficeBatchPlan plan = await OfficeConverter.PlanAsync(settings);
        string expected = Path.Combine(files.Output, expectedName);
        Assert.True(plan.CanRun);
        Assert.Equal(expected, Assert.Single(plan.Items).OutputPath);
        Assert.False(Directory.Exists(files.Output));
        OfficeBatchReport result = await OfficeConverter.ConvertAsync(settings);
        Assert.True(result.Succeeded, Describe(result));
        OfficeBatchItemReport item = Assert.Single(result.Items);
        Assert.Equal(expected, item.OutputPath);
        Assert.Equal(expected, Assert.IsType<OfficeConversionReport>(item.Report).OutputPath);
        OfficeBatchReport restored = OfficeBatchReport.Parse(result.ToJson());
        Assert.Equal(expected, Assert.Single(restored.Items).OutputPath);
        using var pdf = PdfDocument.Load(expected);
        Assert.Contains("Sales document content", pdf.ExtractText());
        Assert.Equal(sourceBytes, File.ReadAllBytes(source));
    }

    [Fact]
    public async Task DirectoryBatchRetainsRelativeFoldersWithTheSameStem() {
        using var files = new NamingFiles();
        files.Write("North/Sales.txt", "North sales");
        files.Write("South/Sales.txt", "South sales");
        var settings = new OfficeBatchOptions { InputDirectory = files.Input, OutputDirectory = files.Output };
        OfficeBatchPlan plan = await OfficeConverter.PlanAsync(settings);
        Assert.True(plan.CanRun); Assert.Equal(0, plan.Collisions);
        OfficeBatchReport report = await OfficeConverter.ConvertDirectoryAsync(files.Input, files.Output, OfficeDocumentFormat.Pdf);
        Assert.True(report.Succeeded, Describe(report)); Assert.Equal(2, report.Summary.Completed);
        foreach (string region in new[] { "North", "South" }) {
            string expected = Path.Combine(files.Output, region, "Sales.pdf");
            Assert.Contains(plan.Items, item => item.OutputPath == expected);
            Assert.Contains(report.Items, item => item.OutputPath == expected);
            using var pdf = PdfDocument.Load(expected);
            Assert.Contains(region + " sales", pdf.ExtractText());
        }
    }

    [Theory]
    [InlineData(OfficePublicationPolicy.Fail)]
    [InlineData(OfficePublicationPolicy.Rename)]
    [InlineData(OfficePublicationPolicy.Replace)]
    public async Task FlattenedSameStemAcrossFormatsUsesTheExistingPublicationPolicy(OfficePublicationPolicy policy) {
        using var files = new NamingFiles();
        string first = files.Write("North/Sales.txt", "First sales");
        string second = files.Write("South/Sales.md", "Second sales");
        OfficeBatchOptions settings = files.Settings(first, second) with { ConflictPolicy = policy, MaxParallelism = 1 };
        OfficeBatchPlan plan = await OfficeConverter.PlanAsync(settings);
        Assert.Equal(1, plan.Collisions); Assert.Equal(policy != OfficePublicationPolicy.Fail, plan.CanRun);
        Assert.All(plan.Items, item => Assert.Equal(Path.Combine(files.Output, "Sales.pdf"), item.OutputPath));
        Assert.False(Directory.Exists(files.Output));
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(settings);
        Assert.Equal(policy == OfficePublicationPolicy.Fail ? 1 : 2, report.Summary.Completed);
        Assert.Equal(policy == OfficePublicationPolicy.Fail ? 1 : 0, report.Summary.Failed);
        Assert.Equal(policy != OfficePublicationPolicy.Fail, report.Succeeded);
        using (var pdf = PdfDocument.Load(Path.Combine(files.Output, "Sales.pdf")))
            Assert.Contains(policy == OfficePublicationPolicy.Replace ? "Second sales" : "First sales", pdf.ExtractText());
        if (policy == OfficePublicationPolicy.Rename) {
            string renamed = Path.Combine(files.Output, "Sales (1).pdf");
            Assert.Contains(report.Items, item => item.OutputPath == renamed);
            using var pdf = PdfDocument.Load(renamed); Assert.Contains("Second sales", pdf.ExtractText());
        }
        Assert.Equal(policy == OfficePublicationPolicy.Rename ? 2 : 1, Directory.GetFiles(files.Output, "*.pdf").Length);
    }

    [Theory]
    [InlineData(OfficePublicationPolicy.Fail)]
    [InlineData(OfficePublicationPolicy.Rename)]
    [InlineData(OfficePublicationPolicy.Replace)]
    public async Task ExistingDestinationIsCheckedAtTheReplacedName(OfficePublicationPolicy policy) {
        using var files = new NamingFiles();
        string source = files.Write("Sales.txt", "New sales");
        Directory.CreateDirectory(files.Output);
        string expected = Path.Combine(files.Output, "Sales.pdf");
        byte[] original = Encoding.UTF8.GetBytes("Keep existing destination"); File.WriteAllBytes(expected, original);
        OfficeBatchOptions settings = files.Settings(source) with { ConflictPolicy = policy };
        OfficeBatchPlan plan = await OfficeConverter.PlanAsync(settings);
        Assert.Equal(1, plan.Collisions); Assert.Equal(policy != OfficePublicationPolicy.Fail, plan.CanRun);
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(settings);
        Assert.Equal(policy != OfficePublicationPolicy.Fail, report.Succeeded);
        if (policy != OfficePublicationPolicy.Replace) Assert.Equal(original, File.ReadAllBytes(expected));
        if (policy != OfficePublicationPolicy.Fail) {
            string output = Assert.Single(report.Items).OutputPath!;
            Assert.Equal(policy == OfficePublicationPolicy.Rename ? Path.Combine(files.Output, "Sales (1).pdf") : expected, output);
            using var pdf = PdfDocument.Load(output); Assert.Contains("New sales", pdf.ExtractText());
        }
    }

    [Fact]
    public async Task ParallelRenamePublishesDistinctArtifactsForFlattenedInputs() {
        using var files = new NamingFiles();
        string first = files.Write("North/Sales.txt", "North concurrent content");
        string second = files.Write("South/Sales.txt", "South concurrent content");
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(files.Settings(first, second) with {
            MaxParallelism = 2, ConflictPolicy = OfficePublicationPolicy.Rename
        });
        Assert.True(report.Succeeded, Describe(report)); Assert.Equal(2, report.Summary.Completed);
        Assert.Equal(new[] { "Sales (1).pdf", "Sales.pdf" }, report.Items.Select(item => Path.GetFileName(item.OutputPath)).OrderBy(name => name, StringComparer.Ordinal));
        foreach (OfficeBatchItemReport item in report.Items) {
            using var pdf = PdfDocument.Load(item.OutputPath!);
            Assert.Contains(File.ReadAllText(item.InputPath), pdf.ExtractText());
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public async Task DurableSelectionCannotShareACompletionArtifactAcrossDistinctInputs(int parallelism) {
        using var files = new NamingFiles();
        // Identical bytes do not turn two distinct selected inputs into one durable destination owner.
        string first = files.Write("North/Sales.txt", "Identical sales content");
        string second = files.Write("South/Sales.txt", "Identical sales content");
        OfficeBatchOptions settings = files.Settings(first, second) with {
            MaxParallelism = parallelism, ResumeFrom = Path.Combine(files.Output, "batch.json")
        };
        OfficeBatchPlan plan = await OfficeConverter.PlanAsync(settings);
        Assert.False(plan.CanRun); Assert.Equal(1, plan.Collisions); Assert.Equal(1, plan.Failed);
        OfficeBatchReport initial = await OfficeConverter.ConvertAsync(settings);
        Assert.False(initial.Succeeded); Assert.Equal(1, initial.Summary.Completed); Assert.Equal(1, initial.Summary.Failed);
        Assert.Equal(0, initial.Summary.Reused);
        string winner = Assert.Single(initial.Items, item => item.Status == OfficeBatchItemStatus.Completed).InputPath;
        string receipt = Assert.Single(Directory.GetFiles(settings.ResumeFrom + ".state/items", "*.json", SearchOption.AllDirectories));
        byte[] receiptBytes = File.ReadAllBytes(receipt), outputBytes = File.ReadAllBytes(Path.Combine(files.Output, "Sales.pdf"));
        // Put the actual receipt owner first for a deterministic reuse assertion. The collision still fails.
        OfficeBatchReport resumed = await OfficeConverter.ConvertAsync(settings with { InputPaths = [winner, winner == first ? second : first], MaxParallelism = 1 });
        Assert.False(resumed.Succeeded); Assert.Equal(1, resumed.Summary.Completed); Assert.Equal(1, resumed.Summary.Reused); Assert.Equal(1, resumed.Summary.Failed);
        Assert.Equal(receiptBytes, File.ReadAllBytes(receipt)); Assert.Equal(outputBytes, File.ReadAllBytes(Path.Combine(files.Output, "Sales.pdf")));
        Assert.Single(Directory.GetFiles(settings.ResumeFrom + ".state/items", "*.json", SearchOption.AllDirectories));
        Assert.Empty(Directory.GetFiles(files.Output, ".*.conversion.pdf"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DuplicatePlanDestinationsStayBlockedWhenPublicationCreatesTheirParent(bool createOutput) {
        using var files = new NamingFiles();
        string first = files.Write("North/Sales.txt", "First source");
        string second = files.Write("South/Sales.txt", "Second source");
        string output = Path.Combine(files.Output, "Sales.pdf");
        var settings = files.Settings(first, second) with { ResumeFrom = Path.Combine(files.Output, "batch.json") };
        Assert.False(Directory.Exists(files.Output));
        var progress = new CreatePublicationEntries(files.Output, output, createOutput);
        OfficeBatchPlan plan = await OfficeConverter.PlanAsync(settings, progress);
        Assert.False(plan.CanRun);
        Assert.Equal(2, plan.Selected);
        Assert.Equal(1, plan.Collisions);
        Assert.Equal(1, plan.Failed);
        Assert.Equal(new[] { "Selected", "Blocked" }, plan.Items.Select(item => item.Disposition));
        Assert.False(File.Exists(settings.ResumeFrom));
        Assert.False(Directory.Exists(settings.ResumeFrom + ".state"));
    }

    private sealed class CreatePublicationEntries(string directory, string output, bool createOutput) : IProgress<OfficeBatchPlanItem> {
        private bool _created;
        public void Report(OfficeBatchPlanItem item) {
            if (_created) return;
            _created = true;
            Directory.CreateDirectory(directory);
            if (createOutput) File.WriteAllText(output, "Published while the remaining files were being planned");
        }
    }

    [Fact]
    public async Task AppendStyleCheckpointPlanIsRejectedWithoutMovingOrReplacingArtifacts() {
        using var files = new NamingFiles(); string source = files.Write("Sales.txt", "Legacy source");
        Directory.CreateDirectory(files.Output);
        OfficeBatchOptions settings = files.Settings(source) with { ResumeFrom = Path.Combine(files.Output, "batch.json") };
        // Reconstruct the previously persisted v2 append-style plan, including its physical identity keys.
        // Reflection reads the existing internal identity implementation; no product test hook is introduced.
        MethodInfo identity = typeof(OfficeColor).Assembly.GetType("OfficeIMO.Internal.OfficePathIdentity", throwOnError: true)!
            .GetMethod("GetPathIdentityKey", BindingFlags.Static | BindingFlags.NonPublic)!;
        string outputIdentity = (string)identity.Invoke(null, [files.Output])!;
        string oldHash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(string.Join("\n",
            "officeimo.conversion.batch.v2", "selected-files", outputIdentity, ".pdf"))));
        byte[] plan = JsonSerializer.SerializeToUtf8Bytes(new { Schema = "officeimo.conversion.batch.v2", ConfigurationSha256 = oldHash });
        File.WriteAllBytes(settings.ResumeFrom!, plan);
        string legacyOutput = Path.Combine(files.Output, "Sales.txt.pdf"); File.WriteAllText(legacyOutput, "Retain legacy artifact");
        await Assert.ThrowsAsync<InvalidDataException>(() => OfficeConverter.ConvertAsync(settings));
        Assert.Equal(plan, File.ReadAllBytes(settings.ResumeFrom!)); Assert.Equal("Retain legacy artifact", File.ReadAllText(legacyOutput));
        Assert.False(File.Exists(Path.Combine(files.Output, "Sales.pdf")));
        Assert.Empty(Directory.GetFiles(settings.ResumeFrom + ".state", "*.json", SearchOption.AllDirectories));
    }

    private static string Describe(OfficeBatchReport report) => string.Join("; ", report.Items.Select(item => item.Summary));

    private sealed class NamingFiles : IDisposable {
        internal string Root { get; } = Directory.CreateTempSubdirectory("officeimo-batch-naming-").FullName;
        internal string Input => Path.Combine(Root, "input");
        internal string Output => Path.Combine(Root, "output");
        internal string PathInInput(string relative) {
            string path = Path.Combine(Input, relative.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(path)!); return path;
        }
        internal string Write(string relative, string text) { string path = PathInInput(relative); File.WriteAllText(path, text); return path; }
        internal OfficeBatchOptions Settings(params string[] sources) => new() { InputPaths = sources, OutputDirectory = Output };
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }
}

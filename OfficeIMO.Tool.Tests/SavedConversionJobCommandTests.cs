using System.Text.Json;
using System.IO.Compression;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class SavedConversionJobCommandTests {
    [Fact]
    public async Task Save_plan_run_and_resume_use_the_job_directory_without_implicit_execution() {
        using var files = new Files();
        File.WriteAllText(Path.Combine(files.Input, "source.txt"), "Saved job flow marker\nfirst\tsecond");

        var saved = await Run("workflow", "batch", "--input-directory", files.Input, "--output", files.Output,
            "--checkpoint", files.State, "--tab-size", "4", "--save-job", files.Job);

        Assert.Equal((int)OfficeImoToolExitCode.Success, saved.Code);
        Assert.True(File.Exists(files.Job));
        Assert.False(Directory.Exists(files.Output));
        Assert.False(Directory.Exists(files.State));
        OfficeBatchOptions definition = OfficeBatchOptions.Load(files.Job);
        Assert.Equal(4, definition.ConversionOptions.PlainText!.TabSize);
        using (var document = JsonDocument.Parse(File.ReadAllText(files.Job)))
            Assert.False(Path.IsPathFullyQualified(document.RootElement.GetProperty("batch").GetProperty("inputDirectory").GetString()!));
        Assert.Equal(files.Input, definition.InputDirectory);

        var plan = await Run("workflow", "job", "plan", files.Job);
        Assert.Equal((int)OfficeImoToolExitCode.Success, plan.Code);
        using (var evidence = JsonDocument.Parse(plan.Output)) {
            Assert.True(evidence.RootElement.GetProperty("canRun").GetBoolean());
            Assert.Equal(1, evidence.RootElement.GetProperty("selected").GetInt64());
        }
        Assert.False(Directory.Exists(files.Output));
        Assert.False(Directory.Exists(files.State));

        var first = await Run("workflow", "job", "run", files.Job);
        Assert.Equal((int)OfficeImoToolExitCode.Success, first.Code);
        string artifact = Path.Combine(files.Output, "source.pdf");
        byte[] original = File.ReadAllBytes(artifact);
        Assert.Contains("Saved job flow marker", PdfReadDocument.Open(original).ExtractText());

        var resumed = await Run("workflow", "job", "run", files.Job, "--retry-failed", "--format", "ndjson");
        Assert.Equal((int)OfficeImoToolExitCode.Success, resumed.Code);
        string lastLine = resumed.Output.Split('\n', StringSplitOptions.RemoveEmptyEntries)[^1];
        using var summary = JsonDocument.Parse(lastLine);
        Assert.Equal(1, summary.RootElement.GetProperty("summary").GetProperty("reused").GetInt64());
        Assert.Equal(original, File.ReadAllBytes(artifact));
    }

    [Fact]
    public async Task Save_keeps_environment_references_without_reading_them_and_refuses_literal_secrets() {
        using var files = new Files();
        File.WriteAllText(Path.Combine(files.Input, "source.txt"), "Saved reference marker");
        string variable = "OFFICEIMO_MISSING_JOB_PASSWORD_" + Guid.NewGuid().ToString("N");

        var saved = await Run("workflow", "batch", "--input-directory", files.Input, "--output", files.Output,
            "--source-password-env", variable, "--save-job", files.Job);
        Assert.Equal((int)OfficeImoToolExitCode.Success, saved.Code);
        Assert.Equal(variable, OfficeBatchOptions.Load(files.Job).SourcePasswordEnvironmentVariable);
        Assert.False(Directory.Exists(files.Output));

        var executed = await Run("workflow", "job", "run", files.Job);
        Assert.NotEqual((int)OfficeImoToolExitCode.Success, executed.Code);
        Assert.False(Directory.Exists(files.Output));

        string otherJob = Path.Combine(Path.GetDirectoryName(files.Job)!, "literal.json");
        var literal = await Run("workflow", "batch", "--input-directory", files.Input, "--output", files.Output,
            "--source-password", "not-a-persistable-secret", "--save-job", otherJob);
        Assert.NotEqual((int)OfficeImoToolExitCode.Success, literal.Code);
        Assert.False(File.Exists(otherJob));
        Assert.DoesNotContain("not-a-persistable-secret", literal.Output + literal.Error);
    }

    [Theory]
    [InlineData("plan")]
    [InlineData("run")]
    public async Task Unsupported_job_schema_fails_before_creating_output(string command) {
        using var files = new Files();
        File.WriteAllText(files.Job, "{\"schema\":\"officeimo.conversion.job.v999\",\"inputDirectory\":\"../input\",\"outputDirectory\":\"../output\"}");

        var rejected = await Run("workflow", "job", command, files.Job);

        Assert.NotEqual((int)OfficeImoToolExitCode.Success, rejected.Code);
        Assert.False(Directory.Exists(files.Output));
    }

    [Fact]
    public async Task Save_does_not_replace_an_existing_job() {
        using var files = new Files();
        const string original = "retained original job";
        File.WriteAllText(files.Job, original);

        var rejected = await Run("workflow", "batch", "--input-directory", files.Input, "--output", files.Output, "--save-job", files.Job);

        Assert.NotEqual((int)OfficeImoToolExitCode.Success, rejected.Code);
        Assert.Equal(original, File.ReadAllText(files.Job));
        Assert.False(Directory.Exists(files.Output));
    }

    [Fact]
    public async Task Directory_package_job_plan_and_run_use_the_same_registered_admission() {
        using var files = new Files();
        string fixture = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "synthetic", "numbers-unpadded-scalar.numbers");
        string source = Path.Combine(files.Input, "source.numbers");
        ZipFile.ExtractToDirectory(fixture, source);
        var saved = await Run("workflow", "batch", source, "--output", files.Output, "--target", "xlsx", "--save-job", files.Job);
        Assert.Equal((int)OfficeImoToolExitCode.Success, saved.Code);

        var planned = await Run("workflow", "job", "plan", files.Job);
        Assert.Equal((int)OfficeImoToolExitCode.Success, planned.Code);
        using (var plan = JsonDocument.Parse(planned.Output)) {
            Assert.True(plan.RootElement.GetProperty("canRun").GetBoolean());
            Assert.Equal(1, plan.RootElement.GetProperty("selected").GetInt64());
            Assert.Equal(0, plan.RootElement.GetProperty("failed").GetInt64());
            JsonElement item = Assert.Single(plan.RootElement.GetProperty("items").EnumerateArray());
            Assert.Equal(source, item.GetProperty("inputPath").GetString());
        }
        Assert.False(Directory.Exists(files.Output));

        var executed = await Run("workflow", "job", "run", files.Job);
        Assert.Equal((int)OfficeImoToolExitCode.Success, executed.Code);
        using (var report = JsonDocument.Parse(executed.Output)) {
            Assert.True(report.RootElement.GetProperty("succeeded").GetBoolean());
            Assert.Equal(1, report.RootElement.GetProperty("summary").GetProperty("completed").GetInt64());
        }
        using var workbook = OfficeIMO.Excel.ExcelDocument.Load(Path.Combine(files.Output, "source.numbers.xlsx"));
        Assert.Equal(42d, Assert.Single(workbook.Sheets).CellAt(1, 1).GetValue<double>());
    }

    [Fact]
    public async Task Recorded_runs_can_be_inspected_and_retried_without_replacing_prior_evidence() {
        using var files = new Files();
        File.WriteAllText(Path.Combine(files.Input, "good.txt"), "Retained completed document");
        string failedSource = Path.Combine(files.Input, "repair.docx");
        File.WriteAllText(failedSource, "Invalid package for the first attempt");
        var saved = await Run("workflow", "batch", "--input-directory", files.Input, "--output", files.Output,
            "--checkpoint", files.State, "--concurrency", "1", "--save-job", files.Job);
        Assert.Equal((int)OfficeImoToolExitCode.Success, saved.Code);
        string firstPath = Path.Combine(Path.GetDirectoryName(files.Job)!, "first-run.json");
        var first = await Run("workflow", "job", "run", files.Job, "--report", firstPath);
        Assert.Equal((int)OfficeImoToolExitCode.OperationFailed, first.Code);
        var report = OfficeBatchReport.Load(firstPath);
        Assert.Equal(1, report.Summary.Completed);
        Assert.Equal(1, report.Summary.Failed);
        Assert.Equal(2, report.Items.Count);
        byte[] retainedReport = File.ReadAllBytes(firstPath);
        string goodOutput = Path.Combine(files.Output, "good.pdf");
        byte[] retainedPdf = File.ReadAllBytes(goodOutput);
        var inspected = await Run("workflow", "job", "report", firstPath);
        Assert.Equal((int)OfficeImoToolExitCode.Success, inspected.Code);
        Assert.Equal(report.RunId, OfficeBatchReport.Parse(inspected.Output).RunId);
        File.Delete(failedSource);
        using (var repaired = OfficeIMO.Word.WordDocument.Create()) {
            repaired.AddParagraph("The failed document is repaired.");
            repaired.Save(failedSource);
        }
        string secondPath = Path.Combine(Path.GetDirectoryName(files.Job)!, "second-run.json");
        var retry = await Run("workflow", "job", "run", files.Job, "--retry-failed", "--report", secondPath);
        Assert.Equal((int)OfficeImoToolExitCode.Success, retry.Code);
        var retried = OfficeBatchReport.Load(secondPath);
        Assert.True(retried.RetryFailed);
        Assert.Equal(2, retried.Summary.Completed);
        Assert.Equal(1, retried.Summary.Reused);
        Assert.NotEqual(report.RunId, retried.RunId);
        Assert.Equal(retainedReport, File.ReadAllBytes(firstPath));
        Assert.Equal(retainedPdf, File.ReadAllBytes(goodOutput));
    }

    [Fact]
    public async Task Recording_preserves_an_occupied_destination_and_plan_remains_read_only() {
        using var files = new Files();
        File.WriteAllText(Path.Combine(files.Input, "one.txt"), "No execution before recording admission");
        var saved = await Run("workflow", "batch", "--input-directory", files.Input, "--output", files.Output, "--save-job", files.Job);
        Assert.Equal((int)OfficeImoToolExitCode.Success, saved.Code);
        string report = Path.Combine(Path.GetDirectoryName(files.Job)!, "occupied.json");
        File.WriteAllText(report, "Unrelated user file");
        var rejected = await Run("workflow", "job", "run", files.Job, "--report", report);
        Assert.NotEqual((int)OfficeImoToolExitCode.Success, rejected.Code);
        Assert.Equal("Unrelated user file", File.ReadAllText(report));
        Assert.False(Directory.Exists(files.Output));
        var plan = await Run("workflow", "job", "plan", files.Job, "--report", report);
        Assert.NotEqual((int)OfficeImoToolExitCode.Success, plan.Code);
        Assert.False(Directory.Exists(files.Output));
    }

    private static async Task<(int Code, string Output, string Error)> Run(params string[] args) {
        using var output = new MemoryStream();
        using var error = new StringWriter();
        int code = await OfficeImoToolApp.RunAsync(args, Stream.Null, output, error);
        return (code, Encoding.UTF8.GetString(output.ToArray()), error.ToString());
    }

    private sealed class Files : IDisposable {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-saved-job-cli-" + Guid.NewGuid().ToString("N"));
        internal string Input => Path.Combine(_root, "input");
        internal string Output => Path.Combine(_root, "output");
        internal string State => Path.Combine(_root, "state");
        internal string Job => Path.Combine(_root, "jobs", "convert.json");
        internal Files() { Directory.CreateDirectory(Input); Directory.CreateDirectory(Path.GetDirectoryName(Job)!); }
        public void Dispose() => Directory.Delete(_root, true);
    }
}

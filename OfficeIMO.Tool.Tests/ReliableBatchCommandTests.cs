using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Tool.Commands.Workflow;
using OfficeIMO.Workflows;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class ReliableBatchCommandTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Empty_selection_requires_explicit_acceptance(bool unsupportedInput) {
        using var files = new Files();
        if (unsupportedInput) File.WriteAllText(Path.Combine(files.Input, "unsupported.bin"), "unconverted");
        string[] arguments = ["--input-directory", files.Input, "--output", files.Output];

        var refused = await RunBatch(arguments);
        var accepted = await RunBatch([.. arguments, "--allow-empty"]);

        Assert.Equal((int)OfficeImoToolExitCode.OperationFailed, refused.Code);
        Assert.Contains("No files were selected", refused.Error);
        Assert.Equal((int)OfficeImoToolExitCode.Success, accepted.Code);
        using var summary = JsonDocument.Parse(refused.Output);
        Assert.Equal(0, summary.RootElement.GetProperty("summary").GetProperty("selected").GetInt64());
        Assert.Equal(unsupportedInput ? 1 : 0, summary.RootElement.GetProperty("summary").GetProperty("skipped").GetInt64());
        Assert.Empty(Directory.GetFiles(files.Output));
    }

    [Theory]
    [InlineData("json")]
    [InlineData("ndjson")]
    public async Task Structured_results_reconcile_success_failure_and_skips(string format) {
        using var files = new Files();
        File.WriteAllText(Path.Combine(files.Input, "good.txt"), "Reliable batch result marker");
        File.WriteAllText(Path.Combine(files.Input, "broken.docx"), "not an Office package");
        File.WriteAllText(Path.Combine(files.Input, "unsupported.bin"), "unconverted");

        var result = await RunBatch(["--input-directory", files.Input, "--output", files.Output, "--format", format]);

        Assert.Equal((int)OfficeImoToolExitCode.OperationFailed, result.Code);
        JsonElement[] items;
        JsonElement summary;
        if (format == "json") {
            using var report = JsonDocument.Parse(result.Output);
            Assert.Equal("officeimo.conversion.batch.v2", report.RootElement.GetProperty("schema").GetString());
            Assert.False(report.RootElement.GetProperty("succeeded").GetBoolean());
            Assert.False(report.RootElement.GetProperty("itemsTruncated").GetBoolean());
            items = report.RootElement.GetProperty("items").EnumerateArray().Select(item => item.Clone()).ToArray();
            summary = report.RootElement.GetProperty("summary").Clone();
        } else {
            var records = result.Output.Split('\n', StringSplitOptions.RemoveEmptyEntries).Select(line => {
                using var document = JsonDocument.Parse(line);
                return document.RootElement.Clone();
            }).ToArray();
            Assert.Equal(4, records.Length);
            Assert.Equal("summary", records[^1].GetProperty("type").GetString());
            Assert.False(records[^1].GetProperty("succeeded").GetBoolean());
            items = records[..^1].Select(record => record.GetProperty("item")).ToArray();
            summary = records[^1].GetProperty("summary");
        }
        Assert.Equal(3, items.Length);
        Assert.Contains(items, item => item.GetProperty("status").ValueKind == JsonValueKind.String && item.GetProperty("status").GetString() == "Completed");
        Assert.Contains(items, item => item.GetProperty("status").ValueKind == JsonValueKind.String && item.GetProperty("status").GetString() == "Failed");
        Assert.Contains(items, item => item.GetProperty("skipped").GetBoolean());
        Assert.Equal(1, summary.GetProperty("completed").GetInt64());
        Assert.Equal(1, summary.GetProperty("failed").GetInt64());
        Assert.Equal(1, summary.GetProperty("skipped").GetInt64());
        Assert.Contains("Reliable batch result marker", PdfDocument.Load(new System.IO.MemoryStream(File.ReadAllBytes(Path.Combine(files.Output, "good.pdf")))).ExtractText());
        Assert.False(File.Exists(Path.Combine(files.Output, "broken.pdf")));
        Assert.Contains("Skipped 1 file:", result.Error, StringComparison.Ordinal);
        Assert.DoesNotContain("unsupported.bin", result.Error, StringComparison.Ordinal);
        Assert.Contains("broken.docx", result.Error, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("json", false)]
    [InlineData("ndjson", false)]
    [InlineData("json", true)]
    [InlineData("ndjson", true)]
    public async Task Discovery_limit_preserves_completed_artifact_and_final_report(string format, bool savedJob) {
        using var files = new Files();
        File.WriteAllText(Path.Combine(files.Input, "first.txt"), "Partial discovery result first");
        File.WriteAllText(Path.Combine(files.Input, "second.txt"), "Partial discovery result second");
        string[] batchArguments = ["--input-directory", files.Input, "--output", files.Output,
            "--maximum-files", "1", "--concurrency", "1"];
        (int Code, string Output, string Error) result;
        if (savedJob) {
            var saved = await RunBatch([.. batchArguments, "--save-job", files.Job]);
            Assert.Equal((int)OfficeImoToolExitCode.Success, saved.Code);
            result = await RunTool(["workflow", "job", "run", files.Job, "--format", format]);
        } else result = await RunBatch([.. batchArguments, "--format", format]);

        Assert.Equal((int)OfficeImoToolExitCode.OperationFailed, result.Code);
        JsonElement summary;
        JsonElement[] items;
        if (format == "json") {
            using var report = JsonDocument.Parse(result.Output);
            Assert.False(report.RootElement.GetProperty("succeeded").GetBoolean());
            summary = report.RootElement.GetProperty("summary").Clone();
            items = report.RootElement.GetProperty("items").EnumerateArray().Select(item => item.Clone()).ToArray();
        } else {
            JsonElement[] records = result.Output.Split('\n', StringSplitOptions.RemoveEmptyEntries).Select(line => {
                using var record = JsonDocument.Parse(line);
                return record.RootElement.Clone();
            }).ToArray();
            Assert.Equal(2, records.Length);
            Assert.Equal("summary", records[^1].GetProperty("type").GetString());
            Assert.False(records[^1].GetProperty("succeeded").GetBoolean());
            summary = records[^1].GetProperty("summary");
            items = records[..^1].Select(record => record.GetProperty("item")).ToArray();
        }
        JsonElement completed = Assert.Single(items);
        Assert.Equal("Completed", completed.GetProperty("status").GetString());
        Assert.Equal(1, summary.GetProperty("selected").GetInt64());
        Assert.Equal(1, summary.GetProperty("completed").GetInt64());
        Assert.Equal(0, summary.GetProperty("failed").GetInt64());
        Assert.False(string.IsNullOrWhiteSpace(summary.GetProperty("failureSummary").GetString()));
        Assert.Contains("Batch stopped:", result.Error);
        string artifact = Assert.Single(Directory.GetFiles(files.Output, "*.pdf"));
        Assert.Equal(artifact, completed.GetProperty("outputPath").GetString());
        Assert.Contains("Partial discovery result", PdfDocument.Load(new System.IO.MemoryStream(File.ReadAllBytes(artifact))).ExtractText());
    }

    [Fact]
    public async Task Batch_uses_registered_iWork_routes() {
        using var files = new Files();
        string fixture = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "synthetic", "numbers-unpadded-scalar.numbers");
        string source = Path.Combine(files.Input, "source.numbers");
        File.Copy(fixture, source);

        var result = await RunBatch([source, "--output", files.Output, "--target", "xlsx"]);

        Assert.Equal((int)OfficeImoToolExitCode.Success, result.Code);
        using var workbook = OfficeIMO.Excel.ExcelDocument.Load(Path.Combine(files.Output, "source.xlsx"));
        Assert.Equal(42d, Assert.Single(workbook.Sheets).Cell(1, 1).GetValue<double>());
    }

    [Theory]
    [InlineData("--source-password")]
    [InlineData("--pdf-password")]
    public async Task Passwords_come_only_from_environment_variables(string literalOption) {
        using var files = new Files();
        string variable = "OFFICEIMO_TEST_MISSING_" + Guid.NewGuid().ToString("N");
        string[] arguments = ["--input-directory", files.Input, "--output", files.Output];

        var missing = await RunBatch([.. arguments, literalOption + "-env", variable]);
        var literal = await RunBatch([.. arguments, literalOption, "never-print-this-secret"]);

        Assert.Equal((int)OfficeImoToolExitCode.Usage, missing.Code);
        Assert.Contains("missing or empty", missing.Error);
        Assert.Equal((int)OfficeImoToolExitCode.Usage, literal.Code);
        Assert.Contains(literalOption + "-env", literal.Error, StringComparison.Ordinal);
        Assert.DoesNotContain("never-print-this-secret", literal.Error + literal.Output);
        Assert.False(Directory.Exists(files.Output));
    }

    [Fact]
    public async Task Named_route_supplies_its_own_target() {
        using var files = new Files();
        string source = Path.Combine(files.Input, "settings.json");
        File.WriteAllText(source, "{\"name\":\"route target marker\"}");

        var result = await RunBatch([source, "--output", files.Output, "--route", "json-yaml"]);

        Assert.Equal((int)OfficeImoToolExitCode.Success, result.Code);
        Assert.Equal("name: route target marker", File.ReadAllText(Path.Combine(files.Output, "settings.yaml")).Trim());
    }

    private static Task<(int Code, string Output, string Error)> RunBatch(string[] args) => RunTool(["workflow", "batch", .. args]);

    private static async Task<(int Code, string Output, string Error)> RunTool(string[] args) {
        using var output = new MemoryStream();
        using var error = new StringWriter();
        int code = await OfficeImoToolApp.RunAsync(args, Stream.Null, output, error);
        return (code, Encoding.UTF8.GetString(output.ToArray()), error.ToString());
    }

    private sealed class Files : IDisposable {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-reliable-cli-" + Guid.NewGuid().ToString("N"));
        internal string Input => Path.Combine(_root, "input");
        internal string Output => Path.Combine(_root, "output");
        internal string Job => Path.Combine(_root, "convert.json");
        internal Files() => Directory.CreateDirectory(Input);
        public void Dispose() => Directory.Delete(_root, true);
    }
}

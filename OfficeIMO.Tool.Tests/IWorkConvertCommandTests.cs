using System.Text.Json;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class IWorkConvertCommandTests {
    [Theory]
    [InlineData("nim-iwork/simple.pages", "docx")]
    [InlineData("nim-iwork/simple.numbers", "xlsx")]
    [InlineData("keynotekit/tabledeck-v15.2.1.key", "pptx")]
    public async Task Apple_default_destinations_publish_OOXML_and_serialize_typed_fidelity(string fixture, string target) {
        using var files = new Files(fixture);
        var result = await RunAsync("convert", files.Input);
        Assert.Equal(0, result.Code);
        using JsonDocument json = JsonDocument.Parse(result.Output);
        Assert.True(json.RootElement.GetProperty("succeeded").GetBoolean());
        string destination = json.RootElement.GetProperty("outputPath").GetString()!;
        Assert.Equal("." + target, Path.GetExtension(destination));
        Assert.True(File.Exists(destination));
        Assert.Contains(json.RootElement.GetProperty("conversionEvidence").GetProperty("fidelityDiagnostics").EnumerateArray(),
            diagnostic => diagnostic.GetProperty("lossKind").GetString() == "Unassessed");
        Assert.Contains(json.RootElement.GetProperty("diagnostics").EnumerateArray(), diagnostic => diagnostic.GetProperty("code").GetString() == "OutputReopened");
    }

    [Fact]
    public async Task Numbers_CLI_retains_source_formula_and_cache_assessments() {
        using var files = new Files("numbers-parser/test-10-formulas.numbers");
        var result = await RunAsync("convert", files.Input);
        Assert.Equal(0, result.Code);
        using JsonDocument json = JsonDocument.Parse(result.Output);
        JsonElement facts = json.RootElement.GetProperty("conversionEvidence").GetProperty("facts");
        Assert.Equal("28", facts.GetProperty("sourceFormulaCellCount").GetString());
        Assert.Equal("28", facts.GetProperty("sourceCompleteFormulaExpressionCount").GetString());
        Assert.Equal("26", facts.GetProperty("sourceCompleteFormulaCacheCount").GetString());
        Assert.Equal("2", facts.GetProperty("sourceApproximateFormulaCacheCount").GetString());
        using var saved = OfficeIMO.Excel.ExcelDocument.Load(json.RootElement.GetProperty("outputPath").GetString()!);
        Assert.Equal(2, saved.Sheets.Count);
    }

    [Fact]
    public async Task Apple_conflicts_require_force_and_replacement_passes_reopen_validation() {
        using var files = new Files("nim-iwork/simple.numbers");
        string destination = Path.ChangeExtension(files.Input, ".xlsx");
        File.WriteAllText(destination, "keep me");
        Assert.Equal((int)OfficeImoToolExitCode.OutputFailed, (await RunAsync("convert", files.Input, destination)).Code);
        Assert.Equal("keep me", File.ReadAllText(destination));
        Assert.Equal(0, (await RunAsync("convert", files.Input, destination, "--force")).Code);
        using var document = OfficeIMO.Excel.ExcelDocument.Load(destination);
        Assert.Single(document.Sheets);
    }

    [Theory]
    [InlineData("--max-input-bytes", "1")]
    [InlineData("--max-output-bytes", "32")]
    [InlineData("--iwork-mode", "visual")]
    public async Task Apple_limits_and_incomplete_visual_coverage_do_not_publish(string option, string value) {
        using var files = new Files("nim-iwork/simple.pages");
        var result = await RunAsync("convert", files.Input, option, value);
        Assert.NotEqual(0, result.Code);
        Assert.False(File.Exists(Path.ChangeExtension(files.Input, ".docx")));
    }

    [Theory]
    [InlineData("--iwork-mode", "invalid")]
    [InlineData("--max-output-bytes", "-1")]
    [InlineData("--max-characters-in-part", "100")]
    public async Task Apple_rejects_invalid_or_unrelated_settings(string option, string value) {
        using var files = new Files("nim-iwork/simple.pages");
        Assert.Equal((int)OfficeImoToolExitCode.Usage, (await RunAsync("convert", files.Input, option, value)).Code);
    }

    [Fact]
    public async Task Explicit_acceptance_of_incomplete_preview_retains_omission_evidence() {
        using var files = new Files("nim-iwork/simple.pages");
        var result = await RunAsync("convert", files.Input, "--iwork-mode", "visual", "--allow-incomplete-preview");
        Assert.Equal(0, result.Code);
        using JsonDocument json = JsonDocument.Parse(result.Output);
        JsonElement evidence = json.RootElement.GetProperty("conversionEvidence");
        Assert.Equal("VisualFallback", evidence.GetProperty("facts").GetProperty("projectionKind").GetString());
        Assert.Contains(evidence.GetProperty("fidelityDiagnostics").EnumerateArray(),
            diagnostic => diagnostic.GetProperty("lossKind").GetString() == "Omission");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Directory_bundle_uses_the_same_CLI_route_and_reports_transport_identity(bool trailingSeparator, bool explicitDestination) {
        using var files = new Files("nim-iwork/simple.numbers");
        string archive = files.Input + ".zip";
        File.Move(files.Input, archive);
        ZipFile.ExtractToDirectory(archive, files.Input);
        string input = files.Input + (trailingSeparator ? Path.DirectorySeparatorChar : string.Empty);
        string destination = Path.ChangeExtension(files.Input, ".xlsx");
        var result = explicitDestination ? await RunAsync("convert", input, destination)
            : await RunAsync("convert", input);
        Assert.Equal(0, result.Code);
        using var json = JsonDocument.Parse(result.Output);
        Assert.Contains(json.RootElement.GetProperty("diagnostics").EnumerateArray(), item =>
            item.GetProperty("code").GetString() == "SourceSnapshot" && item.GetProperty("details").GetProperty("snapshotKind").GetString() == "DirectoryPackage");
        using var output = OfficeIMO.Excel.ExcelDocument.Load(Path.ChangeExtension(files.Input, ".xlsx"));
        Assert.Single(output.Sheets);
    }

    private static async Task<(int Code, string Output, string Error)> RunAsync(params string[] args) {
        using var output = new MemoryStream();
        using var error = new StringWriter();
        int code = await OfficeImoToolApp.RunAsync(args, Stream.Null, output, error);
        return (code, Encoding.UTF8.GetString(output.ToArray()), error.ToString());
    }

    private sealed class Files : IDisposable {
        private readonly string _directory = Path.Combine(Path.GetTempPath(), "officeimo-iwork-cli-" + Guid.NewGuid().ToString("N"));
        public Files(string fixture) {
            Directory.CreateDirectory(_directory);
            Input = Path.Combine(_directory, Path.GetFileName(fixture));
            File.Copy(Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", fixture), Input);
        }
        public string Input { get; }
        public void Dispose() => Directory.Delete(_directory, recursive: true);
    }
}

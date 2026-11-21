using System.Text;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class YamlCommandTests : IDisposable {
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "officeimo-yaml-" + Guid.NewGuid().ToString("N"));

    public YamlCommandTests() => Directory.CreateDirectory(_directory);

    public void Dispose() {
        if (Directory.Exists(_directory)) Directory.Delete(_directory, recursive: true);
    }

    private static async Task<(int ExitCode, string Output, string Error)> RunAsync(string? standardInput, params string[] args) {
        await using var input = new MemoryStream(Encoding.UTF8.GetBytes(standardInput ?? string.Empty));
        await using var output = new MemoryStream();
        using var error = new StringWriter();
        int exitCode = await OfficeImoToolApp.RunAsync(args, input, output, error);
        return (exitCode, Encoding.UTF8.GetString(output.ToArray()), error.ToString());
    }

    private string File(string name, string text, bool byteOrderMark = false) {
        // Names use '/', but directory enumeration reports the platform separator, so expected paths must too.
        string path = Path.Combine(_directory, name.Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        System.IO.File.WriteAllText(path, text, new UTF8Encoding(byteOrderMark));
        return path;
    }

    [Fact]
    public async Task LintReportsFindingsAndFailsOnlyOnErrorsUnlessStrict() {
        string warnings = File("warnings.yaml", "enabled: yes\n");
        string errors = File("nested/errors.yml", "a: 1\na: 2\n");
        File("ignored.txt", "a: [\n");

        (int exitCode, string output, _) = await RunAsync(null, "yaml", "lint", _directory);
        Assert.Equal((int)OfficeImoToolExitCode.ValidationFailed, exitCode);
        Assert.Contains(errors + ":2:1: error ", output, StringComparison.Ordinal);
        Assert.Contains(warnings + ":1:10: warning ", output, StringComparison.Ordinal);
        Assert.DoesNotContain("ignored.txt", output, StringComparison.Ordinal);

        Assert.Equal((int)OfficeImoToolExitCode.Success, (await RunAsync(null, "yaml", "lint", warnings)).ExitCode);
        Assert.Equal((int)OfficeImoToolExitCode.ValidationFailed, (await RunAsync(null, "yaml", "lint", warnings, "--strict")).ExitCode);
        Assert.Equal((int)OfficeImoToolExitCode.Success, (await RunAsync(null, "yaml", "lint", warnings, "--strict", "--disable", "truthy")).ExitCode);
    }

    [Fact]
    public async Task LintJsonCarriesTheSchemaAndFindings() {
        string path = File("a.yaml", "a: 1   \n");
        (int exitCode, string output, _) = await RunAsync(null, "yaml", "lint", path, "--format", "json");
        Assert.Equal(0, exitCode);
        using JsonDocument json = JsonDocument.Parse(output);
        Assert.Equal("officeimo.yaml.lint.v1", json.RootElement.GetProperty("schema").GetString());
        JsonElement finding = json.RootElement.GetProperty("files")[0].GetProperty("findings")[0];
        Assert.Equal("trailing-spaces", finding.GetProperty("rule").GetString());
        Assert.Equal(5, finding.GetProperty("column").GetInt32());
        Assert.True(finding.GetProperty("fixable").GetBoolean());
    }

    [Fact]
    public async Task LintFixRewritesOnlyRepairedLinesAndKeepsTheByteOrderMark() {
        string path = File("fix.yaml", "name:   web   # spacing kept\nenabled: on\nmode: 0755\n", byteOrderMark: true);
        (int exitCode, string output, _) = await RunAsync(null, "yaml", "lint", path, "--fix");
        Assert.Equal(0, exitCode);
        byte[] bytes = System.IO.File.ReadAllBytes(path);
        Assert.Equal(new byte[] { 0xEF, 0xBB, 0xBF }, bytes.Take(3));
        Assert.Equal("name:   web   # spacing kept\nenabled: 'on'\nmode: 0755\n", Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3));
        Assert.Contains(":3:7: warning ", output, StringComparison.Ordinal);
    }

    [Fact]
    public async Task FormatPrintsChecksAndWrites() {
        string path = File("deploy.yaml", "spec:\n    replicas: 2\n    ports:\n    - 80\n");
        (int exitCode, string output, _) = await RunAsync(null, "yaml", "format", path, "--indent", "2", "--sequences", "indented");
        Assert.Equal(0, exitCode);
        Assert.Equal("spec:\n  replicas: 2\n  ports:\n    - 80\n", output);

        (exitCode, output, _) = await RunAsync(null, "yaml", "format", _directory, "--check", "--indent", "2");
        Assert.Equal((int)OfficeImoToolExitCode.ValidationFailed, exitCode);
        Assert.Equal(path, output.Trim());

        Assert.Equal(0, (await RunAsync(null, "yaml", "format", _directory, "--write", "--indent", "2")).ExitCode);
        Assert.Equal("spec:\n  replicas: 2\n  ports:\n  - 80\n", System.IO.File.ReadAllText(path));
        Assert.Equal(0, (await RunAsync(null, "yaml", "format", _directory, "--check", "--indent", "2")).ExitCode);
    }

    [Fact]
    public async Task FormatReadsStandardInput() {
        (int exitCode, string output, _) = await RunAsync("a:   1\nb:  'x'\n", "yaml", "format", "-", "--quotes", "double");
        Assert.Equal(0, exitCode);
        Assert.Equal("a: 1\nb: \"x\"\n", output);
    }

    [Fact]
    public async Task InvalidInputAndUsageReturnTheirExitCodes() {
        string broken = File("broken.yaml", "a: [1\n");
        (int exitCode, _, string error) = await RunAsync(null, "yaml", "format", broken);
        Assert.Equal((int)OfficeImoToolExitCode.UnsupportedInput, exitCode);
        Assert.Contains("(syntax)", error, StringComparison.Ordinal);
        Assert.Equal((int)OfficeImoToolExitCode.ValidationFailed, (await RunAsync(null, "yaml", "lint", broken)).ExitCode);
        Assert.Equal((int)OfficeImoToolExitCode.InputNotFound, (await RunAsync(null, "yaml", "lint", Path.Combine(_directory, "missing.yaml"))).ExitCode);
        Assert.Equal((int)OfficeImoToolExitCode.Usage, (await RunAsync(null, "yaml", "lint", broken, "--enable", "colons")).ExitCode);
        Assert.Equal((int)OfficeImoToolExitCode.Usage, (await RunAsync(null, "yaml", "format", _directory)).ExitCode);
        Assert.Equal((int)OfficeImoToolExitCode.Usage, (await RunAsync("a: 1\n", "yaml", "format", "-", "--write")).ExitCode);
    }

    [Fact]
    public async Task RewritesKeepTheFilePermissions() {
        if (OperatingSystem.IsWindows()) return;
        string path = File("private.yaml", "a:    1   \n");
        System.IO.File.SetUnixFileMode(path, UnixFileMode.UserRead | UnixFileMode.UserWrite);

        Assert.Equal(0, (await RunAsync(null, "yaml", "lint", path, "--fix")).ExitCode);
        Assert.Equal("a:    1\n", System.IO.File.ReadAllText(path));
        Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, System.IO.File.GetUnixFileMode(path));

        Assert.Equal(0, (await RunAsync(null, "yaml", "format", path, "--write")).ExitCode);
        Assert.Equal("a: 1\n", System.IO.File.ReadAllText(path));
        Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, System.IO.File.GetUnixFileMode(path));
    }
}

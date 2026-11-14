using System.Text;
using System.Text.Json;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using OfficeIMO.Excel;
using OfficeIMO.Tool.Agent;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class AccessibilityCommandTests {
    [Fact]
    public async Task CliReportsJsonSarifAndTextWithPolicyExitCodes() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-tool-accessibility-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string warnings = CreateWorkbook(directory, "warnings.xlsx", withHeaderlessTable: false);
            string errors = CreateWorkbook(directory, "errors.xlsx", withHeaderlessTable: true);

            (int exitCode, string output) json = await RunAsync("accessibility", warnings);
            Assert.Equal(0, json.exitCode);
            using (JsonDocument report = JsonDocument.Parse(json.output)) {
                Assert.Equal("Excel", report.RootElement.GetProperty("format").GetString());
                Assert.Equal(0, report.RootElement.GetProperty("errorCount").GetInt32());
                Assert.Contains(report.RootElement.GetProperty("findings").EnumerateArray(),
                    finding => finding.GetProperty("code").GetString() == OfficeDiagnosticCodes.AccessibilityMergedCells);
            }

            Assert.Equal(1, (await RunAsync("accessibility", warnings, "--fail-on-warnings")).ExitCode);
            Assert.Equal(1, (await RunAsync("accessibility", warnings, "--strict")).ExitCode);
            Assert.Equal(1, (await RunAsync("accessibility", errors)).ExitCode);

            string sarifPath = Path.Combine(directory, "out", "errors.sarif");
            Assert.Equal(1, (await RunAsync("accessibility", errors, "--format", "sarif", "--output", sarifPath)).ExitCode);
            using (JsonDocument sarif = JsonDocument.Parse(File.ReadAllText(sarifPath))) {
                Assert.Equal("2.1.0", sarif.RootElement.GetProperty("version").GetString());
                Assert.Contains(sarif.RootElement.GetProperty("runs")[0].GetProperty("results").EnumerateArray(),
                    result => result.GetProperty("ruleId").GetString() == OfficeDiagnosticCodes.AccessibilityMissingTableHeader &&
                        result.GetProperty("level").GetString() == "error");
            }

            (int exitCode, string output) text = await RunAsync("accessibility", errors, "--format", "text");
            Assert.StartsWith("Excel | Default | 1 error(s)", text.output, StringComparison.Ordinal);
            Assert.Contains("Error | " + OfficeDiagnosticCodes.AccessibilityMissingTableHeader + " | Data!A1:B2 (Scores)", text.output, StringComparison.Ordinal);

            Assert.Equal(2, (await RunAsync("accessibility", errors, "--format", "xml")).ExitCode);
            Assert.Equal(3, (await RunAsync("accessibility", Path.Combine(directory, "missing.xlsx"))).ExitCode);
            string notes = Path.Combine(directory, "notes.txt");
            File.WriteAllText(notes, "plain");
            Assert.Equal(4, (await RunAsync("accessibility", notes)).ExitCode);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task McpToolReturnsBoundedFindingsAndOptionalSarif() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-tool-accessibility-mcp-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = CreateWorkbook(directory, "errors.xlsx", withHeaderlessTable: true);
            var transport = new StdioClientTransport(new() {
                Name = "officeimo-accessibility-test", Command = "dotnet", Arguments = [typeof(OfficeImoToolApp).Assembly.Location, "mcp", "serve", "--stdio"],
                WorkingDirectory = directory,
                EnvironmentVariables = new Dictionary<string, string?> { [AgentPathPolicy.AllowedRootsEnvironmentVariable] = directory }
            });
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            await using McpClient client = await McpClient.CreateAsync(transport, cancellationToken: timeout.Token);

            CallToolResult result = await client.CallToolAsync("officeimo_check_accessibility",
                new Dictionary<string, object?> { ["path"] = input, ["maxFindings"] = 1, ["includeSarif"] = true }, cancellationToken: timeout.Token);

            Assert.False(result.IsError);
            JsonElement content = result.StructuredContent!.Value;
            Assert.Equal("Excel", content.GetProperty("format").GetString());
            Assert.Equal("untrusted", content.GetProperty("contentTrust").GetString());
            Assert.Equal(1, content.GetProperty("errorCount").GetInt32());
            Assert.True(content.GetProperty("findingCount").GetInt32() > 1);
            Assert.Single(content.GetProperty("findings").EnumerateArray());
            Assert.True(content.GetProperty("truncated").GetBoolean());
            using JsonDocument sarif = JsonDocument.Parse(content.GetProperty("sarif").GetString()!);
            Assert.Equal(content.GetProperty("findingCount").GetInt32(), sarif.RootElement.GetProperty("runs")[0].GetProperty("results").GetArrayLength());

            CallToolResult outside = await client.CallToolAsync("officeimo_check_accessibility",
                new Dictionary<string, object?> { ["path"] = Path.Combine(Path.GetTempPath(), "elsewhere.xlsx") }, cancellationToken: timeout.Token);
            Assert.True(outside.IsError);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static string CreateWorkbook(string directory, string name, bool withHeaderlessTable) {
        string path = Path.Combine(directory, name);
        using ExcelDocument document = ExcelDocument.Create(path);
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("Name");
        sheet.Cell(1, 2).SetValue("Score");
        sheet.Cell(2, 1).SetValue("Ada");
        sheet.Cell(2, 2).SetValue(10);
        if (withHeaderlessTable) sheet.AddTable("A1:B2", hasHeader: false, name: "Scores");
        sheet.Cell(4, 1).SetValue("Note");
        sheet.Range("A4:B4").Merge();
        sheet.Cell(3, 1).SetValue("Spacer");
        document.Save();
        return path;
    }

    private static async Task<(int ExitCode, string Output)> RunAsync(params string[] args) {
        await using var output = new MemoryStream();
        using var error = new StringWriter();
        int exitCode = await OfficeImoToolApp.RunAsync(args, Stream.Null, output, error);
        return (exitCode, Encoding.UTF8.GetString(output.ToArray()) + error);
    }
}

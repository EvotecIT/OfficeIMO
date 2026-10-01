using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed partial class ProvenanceCommandTests {
    [Fact]
    public async Task AuditNdjsonAndCheckSarifUseTheSameReadOnlyEvidence() {
        using var scope = new TestDirectory();
        string first = scope.Write("first.txt", "review\u202Ethis"); string second = scope.Write("second.txt", "safe");
        string before = File.ReadAllText(first);
        var audit = await RunAsync(["provenance", "audit", first, second, "--format", "ndjson"]);
        Assert.Equal(0, audit.ExitCode);
        string[] lines = audit.Output.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(2, lines.Length);
        foreach (string line in lines) { using var item = JsonDocument.Parse(line); Assert.Equal("officeimo.provenance.result.v2", item.RootElement.GetProperty("schema").GetString()); }
        var check = await RunAsync(["provenance", "check", first, second, "--format", "sarif"]);
        Assert.Equal(1, check.ExitCode);
        using var sarif = JsonDocument.Parse(check.Output);
        Assert.Equal("2.1.0", sarif.RootElement.GetProperty("version").GetString());
        Assert.Contains(sarif.RootElement.GetProperty("runs")[0].GetProperty("results").EnumerateArray(), item => item.GetProperty("ruleId").GetString() == "officeimo.text.BidirectionalControl");
        Assert.Equal(before, File.ReadAllText(first));
        var failed = await RunAsync(["provenance", "check", Path.Combine(Path.GetDirectoryName(first)!, "missing.txt")]);
        Assert.Equal(3, failed.ExitCode);
    }
    [Fact]
    public async Task DisabledTextInspectionIsExplicitInTextAndJson() {
        using var scope = new TestDirectory(); string input = scope.Write("input.txt", "\u202E");
        var text = await RunAsync(["provenance", "assess", input, "--no-text-integrity", "--format", "text"]);
        Assert.Equal(0, text.ExitCode); Assert.Contains("Text integrity: Disabled", text.Output); Assert.DoesNotContain("Text integrity: 0", text.Output);
        var invalid = await RunAsync(["provenance", "capabilities", "--format", "sarif"]);
        Assert.Equal(2, invalid.ExitCode);
    }
}

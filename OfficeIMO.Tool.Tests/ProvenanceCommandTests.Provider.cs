using System.Text.Json;
using OfficeIMO.Tool.Commands.Provenance;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed partial class ProvenanceCommandTests {
    [Fact]
    public void ProviderOptionsReachAssessmentWithoutEnablingNetworkAccess() {
        var parsed = ProvenanceArguments.Parse(["assess", "input.jpg", "--c2patool", "trusted-tool",
            "--trust-anchors", "anchors.pem", "--allowed-list", "allowed.pem", "--verification-timeout-seconds", "12"]);
        var request = ProvenanceCommand.CreateRequest(parsed, "input.jpg", null);
        Assert.Equal("trusted-tool", parsed.C2paToolPath);
        Assert.Equal("anchors.pem", request.Assessment.Verification.TrustAnchorsPath);
        Assert.Equal("allowed.pem", request.Assessment.Verification.AllowedListPath);
        Assert.Equal(TimeSpan.FromSeconds(12), request.Assessment.Verification.Timeout);
        Assert.False(request.Assessment.Verification.AllowNetworkAccess);
    }
    [Theory]
    [InlineData("inspect", "--c2patool", "tool")]
    [InlineData("assess", "--trust-anchors", "anchors.pem")]
    [InlineData("doctor", "--trust-anchors", "anchors.pem")]
    public void MisappliedProviderOptionsFailInsteadOfBeingIgnored(string command, string option, string value) {
        Assert.Throws<ProvenanceUsageException>(() => ProvenanceArguments.Parse([command, "input.jpg", option, value]));
    }
    [Fact]
    public async Task ExplicitlyUnavailableVerifierDoesNotReturnSuccess() {
        using var scope = new TestDirectory();
        string input = scope.Write("note.txt", "ordinary text");
        var result = await RunAsync(["provenance", "assess", input, "--c2patool", Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"))]);
        Assert.Equal((int)OfficeImoToolExitCode.OperationFailed, result.ExitCode);
        using var json = JsonDocument.Parse(result.Output);
        Assert.Equal("Failed", json.RootElement.GetProperty("checks").GetProperty("verification").GetString());
    }
    [Fact]
    public async Task DoctorReportsMissingToolSeparatelyFromAssetVerification() {
        var result = await RunAsync(["provenance", "doctor", "--c2patool", Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"))]);
        Assert.Equal((int)OfficeImoToolExitCode.OperationFailed, result.ExitCode);
        using var json = JsonDocument.Parse(result.Output);
        Assert.Equal("officeimo.provenance.doctor.v1", json.RootElement.GetProperty("schema").GetString());
        Assert.False(json.RootElement.GetProperty("available").GetBoolean());
        Assert.False(string.IsNullOrWhiteSpace(json.RootElement.GetProperty("diagnostic").GetString()));
    }
}

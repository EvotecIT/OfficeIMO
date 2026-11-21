using System.Text.Json;
using static OfficeIMO.Email.Tests.EmailAuthenticationTestSupport;

namespace OfficeIMO.Email.Tests;

public sealed class EmailArcTests {
    // The suite predates RFC 8617 in two places; the verifier follows the RFC and fails closed:
    // - RFC 8617 section 4.1.2 gives ARC-Message-Signature the DKIM c= default (simple/simple); the suite signed
    //   ams_fields_c_na with relaxed canonicalization and no c= tag.
    // - RFC 6376 section 3.5 h= grammar has no empty header names; ams_fields_h_mis_hdr uses "to::date".
    private static readonly Dictionary<string, string> RfcDivergences = new(StringComparer.Ordinal) {
        ["ams_fields_c_na"] = "Fail",
        ["ams_fields_h_mis_hdr"] = "Fail",
    };

    public static TheoryData<string> SuiteCases() {
        var data = new TheoryData<string>();
        using JsonDocument document = ReadJsonFixture("arc-test-suite-validation.json");
        foreach (JsonElement item in document.RootElement.EnumerateArray()) data.Add(item.GetProperty("name").GetString()!);
        return data;
    }

    [Theory]
    [MemberData(nameof(SuiteCases))]
    public async Task ArcTestSuiteValidationCases(string name) {
        using JsonDocument document = ReadJsonFixture("arc-test-suite-validation.json");
        JsonElement item = document.RootElement.EnumerateArray().Single(entry => entry.GetProperty("name").GetString() == name);
        var resolver = new EmailDkimStaticKeyResolver();
        foreach (JsonProperty record in item.GetProperty("txtRecords").EnumerateObject()) resolver.Add(record.Name, record.Value.GetString()!);
        byte[] message = Encoding.UTF8.GetBytes(item.GetProperty("message").GetString()!);

        EmailArcVerificationResult result = await EmailArc.VerifyAsync(message, resolver, Verifier);

        // Three suite cases leave the expected value blank: the newest seal already records cv=fail, so RFC 8617
        // section 5.2 step 3 makes the chain fail.
        string expected = item.GetProperty("cv").GetString()!.Trim();
        if (expected.Length == 0) expected = "Fail";
        if (RfcDivergences.TryGetValue(name, out string? rfcResult)) expected = rfcResult;
        Assert.Equal(expected, result.Status.ToString());
        if (result.Status == EmailArcChainStatus.Pass) {
            Assert.All(result.Sets, set => Assert.True(set.Seal!.IsPass));
            Assert.True(result.Sets[result.Sets.Count - 1].MessageSignature!.IsPass);
        }
    }

    [Fact]
    public async Task OlderMessageSignaturesMayBreakWithoutFailingTheChain() {
        using JsonDocument document = ReadJsonFixture("arc-test-suite-validation.json");
        JsonElement item = document.RootElement.EnumerateArray().Single(entry => entry.GetProperty("name").GetString() == "cv_pass_i2_1_ams1_invalid");
        var resolver = new EmailDkimStaticKeyResolver();
        foreach (JsonProperty record in item.GetProperty("txtRecords").EnumerateObject()) resolver.Add(record.Name, record.Value.GetString()!);

        EmailArcVerificationResult result = await EmailArc.VerifyAsync(Encoding.UTF8.GetBytes(item.GetProperty("message").GetString()!), resolver, Verifier);

        Assert.Equal(EmailArcChainStatus.Pass, result.Status);
        Assert.Equal(2, result.OldestPassingInstance);
        Assert.False(result.Sets[0].MessageSignature!.IsPass);
        Assert.StartsWith("lists.example.org", result.Sets[1].AuthenticationResults, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MissingKeysMakeTheChainNotVerifiedRatherThanFailed() {
        using JsonDocument document = ReadJsonFixture("arc-test-suite-validation.json");
        JsonElement item = document.RootElement.EnumerateArray().Single(entry => entry.GetProperty("name").GetString() == "cv_pass_i1_1");

        EmailArcVerificationResult result = await EmailArc.VerifyAsync(Encoding.UTF8.GetBytes(item.GetProperty("message").GetString()!),
            new UnavailableResolver(), Verifier);

        Assert.Equal(EmailArcChainStatus.NotVerified, result.Status);
        Assert.Equal(EmailArcFailureReason.NotEvaluated, result.FailureReason);
    }

    private sealed class UnavailableResolver : IEmailDkimKeyResolver {
        public Task<EmailDkimKeyLookupResult> ResolveTxtAsync(string queryName, CancellationToken cancellationToken) =>
            Task.FromResult(EmailDkimKeyLookupResult.TemporaryFailure("timeout"));
    }
}

using System.Text.Json;
using OfficeIMO.Security;
using static OfficeIMO.Email.Tests.EmailAuthenticationTestSupport;

namespace OfficeIMO.Email.Tests;

public sealed class EmailDkimTests {
    private const string Rfc8463Ed25519Record = "v=DKIM1; k=ed25519; p=11qYAYKxCrfVS/7TyWQHOg7hcvPapiMlrwIaaPcHURo=";
    private const string Rfc8463RsaRecord = "v=DKIM1; k=rsa; p=MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQDkHlOQoBTzWR" +
        "iGs5V6NpP3idY6Wk08a5qhdR6wy5bdOKb2jLQiY/J16JYi0Qvx/byYzCNb3W91y3FutAC" +
        "DfzwQ/BC/e/8uBsCR+yz1Lxj+PL6lHvqMKrM3rG4hstT5QjvHO9PzoxZyVYLzBfO2EeC3" +
        "Ip3G+2kryOTIKT+l/K4w3QIDAQAB";

    private static readonly byte[] PlainMessage = Ascii(
        "From: Alice <alice@example.com>\n" +
        "To: Bob <bob@example.net>\n" +
        "Subject: Status\n" +
        "Date: Fri, 13 Nov 2026 10:00:00 +0000\n" +
        "\n" +
        "Body line one.\n" +
        "Body line two.\n");

    private static readonly Lazy<TestDkimSigner> SharedRsaSigner = new(() => TestDkimSigner.Rsa(2048));

    private static EmailDkimStaticKeyResolver Rfc8463Keys() => new EmailDkimStaticKeyResolver()
        .Add("brisbane", "football.example.com", Rfc8463Ed25519Record)
        .Add("test", "football.example.com", Rfc8463RsaRecord);

    [Fact]
    public async Task Rfc8463AppendixMessageVerifiesBothSignatures() {
        EmailDkimVerificationResult result = await EmailDkim.VerifyAsync(ReadFixture("rfc8463-appendix-a.eml"), Rfc8463Keys(), Verifier);

        Assert.Collection(result.Signatures,
            ed25519 => {
                Assert.Equal(EmailDkimResult.Pass, ed25519.Result);
                Assert.Equal("ed25519-sha256", ed25519.Algorithm);
                Assert.Equal(256, ed25519.KeySizeBits);
                Assert.Equal(new[] { "from", "to", "subject", "date", "message-id", "from", "subject", "date" }, ed25519.SignedHeaders);
            },
            rsa => {
                Assert.Equal(EmailDkimResult.Pass, rsa.Result);
                Assert.Equal("rsa-sha256", rsa.Algorithm);
                Assert.Equal(1024, rsa.KeySizeBits);
                Assert.Equal("@football.example.com", rsa.Identity);
            });
    }

    [Fact]
    public async Task Rfc8463AppendixMessageFailsAfterBodyOrHeaderChange() {
        string original = Encoding.ASCII.GetString(ReadFixture("rfc8463-appendix-a.eml"));
        byte[] body = Encoding.ASCII.GetBytes(original.Replace("We lost the game.", "We won the game."));
        byte[] header = Encoding.ASCII.GetBytes(original.Replace("Subject: Is dinner ready?", "Subject: Is lunch ready?"));

        EmailDkimVerificationResult bodyResult = await EmailDkim.VerifyAsync(body, Rfc8463Keys(), Verifier);
        EmailDkimVerificationResult headerResult = await EmailDkim.VerifyAsync(header, Rfc8463Keys(), Verifier);

        Assert.All(bodyResult.Signatures, signature => Assert.Equal(EmailDkimFailureReason.BodyHashMismatch, signature.FailureReason));
        Assert.All(headerResult.Signatures, signature => {
            Assert.Equal(EmailDkimResult.Fail, signature.Result);
            Assert.Equal(EmailDkimFailureReason.SignatureMismatch, signature.FailureReason);
        });
    }

    [Fact]
    public void Rfc6376CanonicalizationExamplesMatch() {
        // RFC 6376 section 3.4.6.
        Assert.Equal("a:X\r\n", EmailDkimCanonicalizer.CanonicalizeHeader("A: X\r\n", EmailDkimCanonicalization.Relaxed));
        Assert.Equal("b:Y Z\r\n", EmailDkimCanonicalizer.CanonicalizeHeader("B : Y\t\r\n\tZ  \r\n", EmailDkimCanonicalization.Relaxed));
        Assert.Equal("B : Y\t\r\n\tZ  \r\n", EmailDkimCanonicalizer.CanonicalizeHeader("B : Y\t\r\n\tZ  \r\n", EmailDkimCanonicalization.Simple));

        byte[] body = Encoding.ASCII.GetBytes(" C \r\nD \t E\r\n\r\n\r\n");
        Assert.Equal(" C\r\nD E\r\n", Encoding.ASCII.GetString(EmailDkimCanonicalizer.CanonicalizeBody(body, EmailDkimCanonicalization.Relaxed)));
        Assert.Equal(" C \r\nD \t E\r\n", Encoding.ASCII.GetString(EmailDkimCanonicalizer.CanonicalizeBody(body, EmailDkimCanonicalization.Simple)));

        // RFC 6376 sections 3.4.3 and 3.4.4: empty bodies.
        Assert.Equal("\r\n", Encoding.ASCII.GetString(EmailDkimCanonicalizer.CanonicalizeBody(Array.Empty<byte>(), EmailDkimCanonicalization.Simple)));
        Assert.Empty(EmailDkimCanonicalizer.CanonicalizeBody(Encoding.ASCII.GetBytes("\r\n\r\n"), EmailDkimCanonicalization.Relaxed));
    }

    public static TheoryData<string> DkimpyCases() {
        var data = new TheoryData<string>();
        using JsonDocument document = ReadJsonFixture("dkimpy-signed.json");
        foreach (JsonElement item in document.RootElement.GetProperty("cases").EnumerateArray()) data.Add(item.GetProperty("name").GetString()!);
        return data;
    }

    [Theory]
    [MemberData(nameof(DkimpyCases))]
    public async Task MessagesSignedByDkimpyVerifyAsExpected(string name) {
        using JsonDocument document = ReadJsonFixture("dkimpy-signed.json");
        var resolver = new EmailDkimStaticKeyResolver();
        foreach (JsonProperty record in document.RootElement.GetProperty("records").EnumerateObject()) resolver.Add(record.Name, record.Value.GetString()!);
        JsonElement item = document.RootElement.GetProperty("cases").EnumerateArray().Single(entry => entry.GetProperty("name").GetString() == name);

        EmailDkimVerificationResult result = await EmailDkim.VerifyAsync(Convert.FromBase64String(item.GetProperty("message").GetString()!), resolver, Verifier);

        EmailDkimSignatureResult signature = Assert.Single(result.Signatures);
        Assert.Equal(item.GetProperty("expected").GetString(), signature.Result.ToString());
        if (item.GetProperty("reason").ValueKind == JsonValueKind.String) {
            Assert.Equal(item.GetProperty("reason").GetString(), signature.FailureReason.ToString());
        }

        switch (name) {
            case "length_limit_appended":
                Assert.True(signature.UnsignedBodyBytes > 0);
                Assert.Contains(signature.Warnings, warning => warning.Contains("not signed", StringComparison.Ordinal));
                break;
            case "not_oversigned_from_added":
                Assert.Contains(signature.Warnings, warning => warning.Contains("from", StringComparison.OrdinalIgnoreCase));
                break;
            case "bare_lf_storage":
                Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_AUTH_LINE_ENDINGS_NORMALIZED");
                break;
        }
    }

    public static TheoryData<string, string, EmailDkimResult, EmailDkimFailureReason> PolicyCases() => new() {
        { "unknown version", "v=2; a=rsa-sha256; d=example.com; s=k; h=from", EmailDkimResult.PermError, EmailDkimFailureReason.UnsupportedVersion },
        { "duplicate tag", "v=1; a=rsa-sha256; d=example.com; d=example.com; s=k; h=from", EmailDkimResult.PermError, EmailDkimFailureReason.MalformedSignature },
        { "unknown algorithm", "v=1; a=rsa-sha512; d=example.com; s=k; h=from", EmailDkimResult.PermError, EmailDkimFailureReason.UnsupportedAlgorithm },
        { "from not signed", "v=1; a=rsa-sha256; d=example.com; s=k; h=to:subject", EmailDkimResult.PermError, EmailDkimFailureReason.InvalidSignedHeaders },
        { "identity outside domain", "v=1; a=rsa-sha256; d=example.com; i=@example.org; s=k; h=from", EmailDkimResult.PermError, EmailDkimFailureReason.IdentityMismatch },
        { "query method", "v=1; a=rsa-sha256; d=example.com; q=http; s=k; h=from", EmailDkimResult.PermError, EmailDkimFailureReason.UnsupportedQueryMethod },
        { "expired", "v=1; a=rsa-sha256; d=example.com; s=k; h=from; t=1000; x=2000", EmailDkimResult.Fail, EmailDkimFailureReason.Expired },
        { "expiry before timestamp", "v=1; a=rsa-sha256; d=example.com; s=k; h=from; t=2000; x=1000", EmailDkimResult.PermError, EmailDkimFailureReason.InvalidTimestamps },
        { "no key", "v=1; a=rsa-sha256; d=example.com; s=missing; h=from", EmailDkimResult.PermError, EmailDkimFailureReason.KeyNotFound },
        { "temporary failure", "v=1; a=rsa-sha256; d=example.com; s=down; h=from", EmailDkimResult.TempError, EmailDkimFailureReason.KeyLookupFailed },
        { "resolver throws", "v=1; a=rsa-sha256; d=example.com; s=throws; h=from", EmailDkimResult.TempError, EmailDkimFailureReason.KeyLookupFailed },
        { "two key records", "v=1; a=rsa-sha256; d=example.com; s=twice; h=from", EmailDkimResult.PermError, EmailDkimFailureReason.KeyRecordMalformed },
        { "revoked", "v=1; a=rsa-sha256; d=example.com; s=revoked; h=from", EmailDkimResult.PermError, EmailDkimFailureReason.KeyRevoked },
        { "key type", "v=1; a=rsa-sha256; d=example.com; s=edtype; h=from", EmailDkimResult.PermError, EmailDkimFailureReason.KeyTypeMismatch },
        { "key hash", "v=1; a=rsa-sha256; d=example.com; s=sha1only; h=from", EmailDkimResult.PermError, EmailDkimFailureReason.KeyHashNotAllowed },
        { "key service", "v=1; a=rsa-sha256; d=example.com; s=service; h=from", EmailDkimResult.PermError, EmailDkimFailureReason.KeyServiceNotAllowed },
        { "strict identity", "v=1; a=rsa-sha256; d=example.com; i=@mail.example.com; s=strict; h=from", EmailDkimResult.PermError, EmailDkimFailureReason.KeyStrictIdentity },
        { "bad key record", "v=1; a=rsa-sha256; d=example.com; s=garbage; h=from", EmailDkimResult.PermError, EmailDkimFailureReason.KeyRecordMalformed },
        { "subdomain identity", "v=1; a=rsa-sha256; d=example.com; i=user@mail.example.com; s=k; h=from", EmailDkimResult.Pass, EmailDkimFailureReason.None },
    };

    [Theory]
    [MemberData(nameof(PolicyCases))]
    public async Task SignatureAndKeyRecordRulesFailClosed(string _, string tags, EmailDkimResult expected, EmailDkimFailureReason reason) {
        TestDkimSigner signer = SharedRsaSigner.Value;
        string record = signer.KeyRecord;
        var resolver = new PolicyResolver(new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase) {
            ["k._domainkey.example.com"] = new[] { record },
            ["twice._domainkey.example.com"] = new[] { record, record },
            ["revoked._domainkey.example.com"] = new[] { "v=DKIM1; k=rsa; p=" },
            ["edtype._domainkey.example.com"] = new[] { record.Replace("k=rsa", "k=ed25519") },
            ["sha1only._domainkey.example.com"] = new[] { record.Replace("k=rsa;", "k=rsa; h=sha1;") },
            ["service._domainkey.example.com"] = new[] { record + "; s=voice" },
            ["strict._domainkey.example.com"] = new[] { record + "; t=s" },
            ["garbage._domainkey.example.com"] = new[] { "v=DKIM1; k=rsa; this is not a tag list" },
        });

        byte[] message = signer.Sign(PlainMessage, tags);
        EmailDkimVerificationResult result = await EmailDkim.VerifyAsync(message, resolver, Verifier,
            new EmailAuthenticationOptions { VerificationTime = DateTimeOffset.FromUnixTimeSeconds(1_800_000_000) });

        EmailDkimSignatureResult signature = Assert.Single(result.Signatures);
        Assert.Equal(expected, signature.Result);
        Assert.Equal(reason, signature.FailureReason);
    }

    [Fact]
    public async Task SmallRsaKeysAreRefusedByPolicy() {
        TestDkimSigner signer = TestDkimSigner.Rsa(512);
        byte[] message = signer.Sign(PlainMessage, "v=1; a=rsa-sha256; d=example.com; s=k; h=from");
        var resolver = new EmailDkimStaticKeyResolver().Add("k", "example.com", signer.KeyRecord);

        EmailDkimSignatureResult signature = Assert.Single((await EmailDkim.VerifyAsync(message, resolver, Verifier)).Signatures);

        Assert.Equal(EmailDkimResult.Policy, signature.Result);
        Assert.Equal(EmailDkimFailureReason.KeyTooSmall, signature.FailureReason);
        Assert.Equal(512, signature.KeySizeBits);
    }

    [Fact]
    public async Task BodyLengthLimitIsReportedAndCanBeRefused() {
        TestDkimSigner signer = TestDkimSigner.Ed25519("; t=y");
        byte[] message = signer.Sign(PlainMessage, "v=1; a=ed25519-sha256; d=example.com; s=k; h=from; l=16", length: 16);
        byte[] tooLong = signer.Sign(PlainMessage, "v=1; a=ed25519-sha256; d=example.com; s=k; h=from; l=9999", computeBodyHash: false);
        var resolver = new EmailDkimStaticKeyResolver().Add("k", "example.com", signer.KeyRecord);

        EmailDkimSignatureResult warned = Assert.Single((await EmailDkim.VerifyAsync(message, resolver, Verifier)).Signatures);
        EmailDkimSignatureResult refused = Assert.Single((await EmailDkim.VerifyAsync(message, resolver, Verifier,
            new EmailAuthenticationOptions { RejectBodyLengthLimit = true })).Signatures);
        EmailDkimSignatureResult exceeded = Assert.Single((await EmailDkim.VerifyAsync(tooLong, resolver, Verifier)).Signatures);

        Assert.Equal(EmailDkimResult.Pass, warned.Result);
        Assert.Equal(16, warned.BodyLengthLimit);
        Assert.Equal(16, warned.UnsignedBodyBytes);
        Assert.True(warned.IsTestingKey);
        Assert.Equal(EmailDkimResult.Policy, refused.Result);
        Assert.Equal(EmailDkimFailureReason.BodyLengthLimitRefused, refused.FailureReason);
        Assert.Equal(EmailDkimFailureReason.BodyLengthExceedsBody, exceeded.FailureReason);
    }

    [Fact]
    public async Task VerifierWithoutEd25519ReportsNeutral() {
        EmailDkimVerificationResult result = await EmailDkim.VerifyAsync(ReadFixture("rfc8463-appendix-a.eml"), Rfc8463Keys(), new RsaOnlyVerifier());

        Assert.Equal(EmailDkimResult.Neutral, result.Signatures[0].Result);
        Assert.Equal(EmailDkimFailureReason.VerifierUnavailable, result.Signatures[0].FailureReason);
        Assert.Equal(EmailDkimResult.Pass, result.Signatures[1].Result);
    }

    [Fact]
    public async Task DocumentOverloadRequiresRetainedSourceBytes() {
        byte[] raw = ReadFixture("rfc8463-appendix-a.eml");
        EmailDocument preserved = EmailDocument.Load(raw, new EmailReaderOptions(preserveRawSource: true));
        EmailDocument regenerated = EmailDocument.Load(raw);

        EmailDkimVerificationResult verified = await EmailDkim.VerifyAsync(preserved, Rfc8463Keys(), Verifier);
        EmailDkimVerificationResult missing = await EmailDkim.VerifyAsync(regenerated, Rfc8463Keys(), Verifier);

        Assert.Equal(2, verified.Signatures.Count(signature => signature.IsPass));
        Assert.Empty(missing.Signatures);
        Assert.Contains(missing.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_AUTH_RAW_SOURCE_MISSING");
    }

    [Fact]
    public async Task HostileInputIsBoundedAndNeverThrows() {
        var random = new Random(6376);
        var resolver = Rfc8463Keys();
        byte[] vector = ReadFixture("rfc8463-appendix-a.eml");
        for (int iteration = 0; iteration < 300; iteration++) {
            byte[] mutated = (byte[])vector.Clone();
            for (int change = 0; change < 1 + iteration % 8; change++) mutated[random.Next(mutated.Length)] = (byte)random.Next(256);
            EmailDkimVerificationResult result = await EmailDkim.VerifyAsync(mutated, resolver, Verifier);
            Assert.All(result.Signatures, signature => Assert.False(string.IsNullOrEmpty(signature.Message)));
        }

        string many = string.Concat(Enumerable.Repeat("DKIM-Signature: v=1; a=rsa-sha256; d=example.com; s=k; h=from; bh=AA==; b=AA==\r\n", 40));
        EmailDkimVerificationResult limited = await EmailDkim.VerifyAsync(Encoding.ASCII.GetBytes(many + "From: a@example.com\r\n\r\nx\r\n"),
            resolver, Verifier);
        Assert.Equal(16, limited.Signatures.Count);
        Assert.Contains(limited.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_DKIM_SIGNATURE_LIMIT");

        EmailDkimVerificationResult tooLarge = await EmailDkim.VerifyAsync(vector, resolver, Verifier,
            new EmailAuthenticationOptions { MaxMessageBytes = 100 });
        Assert.Empty(tooLarge.Signatures);
        Assert.Contains(tooLarge.Diagnostics, diagnostic => diagnostic.Code == "EMAIL_AUTH_MESSAGE_TOO_LARGE");
    }

    private sealed class PolicyResolver : IEmailDkimKeyResolver {
        private readonly Dictionary<string, string[]> _records;

        internal PolicyResolver(Dictionary<string, string[]> records) => _records = records;

        public Task<EmailDkimKeyLookupResult> ResolveTxtAsync(string queryName, CancellationToken cancellationToken) {
            if (queryName.StartsWith("down.", StringComparison.Ordinal)) return Task.FromResult(EmailDkimKeyLookupResult.TemporaryFailure());
            if (queryName.StartsWith("throws.", StringComparison.Ordinal)) throw new IOException("network down");
            return Task.FromResult(_records.TryGetValue(queryName, out string[]? values)
                ? EmailDkimKeyLookupResult.Found(values)
                : EmailDkimKeyLookupResult.NotFound());
        }
    }

    private sealed class RsaOnlyVerifier : IOfficePublicKeySignatureVerifier {
        public string Name => "rsa-only";

        public bool Supports(OfficePublicKeySignatureAlgorithm algorithm) => algorithm == OfficePublicKeySignatureAlgorithm.RsaPkcs1Sha256;

        public OfficePublicKeySignatureCheck Verify(OfficePublicKeySignatureAlgorithm algorithm, byte[] publicKey, byte[] data, byte[] signature) =>
            OfficeSecurityProvider.Default.Verify(algorithm, publicKey, data, signature);
    }
}

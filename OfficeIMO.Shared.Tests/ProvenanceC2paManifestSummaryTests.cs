using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

/// <summary>Reads what a C2PA manifest says: generator, actions, ingredients, and signer.</summary>
public sealed class ProvenanceC2paManifestSummaryTests {
    private const string TrainedAlgorithmicMedia = "http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia";

    [Fact]
    public void GenerativeImageManifestNamesTheGeneratorModelAndSigner() {
        byte[] store = Store(
            Manifest("urn:uuid:first", claimGenerator: "Photoshop", actions: Actions(("c2pa.opened", null, null)), signer: null),
            Manifest("urn:uuid:active",
                claimGenerator: null,
                claimGeneratorInfo: ("ChatGPT", null),
                actions: Actions(("c2pa.created", "GPT-4o", TrainedAlgorithmicMedia), ("c2pa.converted", null, null)),
                ingredient: "prompt-image.png",
                signer: Certificate(subjectOrganization: "OpenAI", issuerOrganization: "Truepic")));

        OfficeC2paManifestSummary? summary = OfficeC2paManifestStore.TryDescribe(store, 0, store.Length);

        Assert.NotNull(summary);
        Assert.Equal("urn:uuid:active", summary!.Label);
        Assert.Equal(2, summary.ManifestCount);
        Assert.Equal("ChatGPT", summary.ClaimGenerator);
        Assert.Equal("image.png", summary.Title);
        Assert.Equal("image/png", summary.Format);
        Assert.Equal("OpenAI", summary.SignedBy);
        Assert.Equal("Truepic", summary.CertificateIssuer);
        Assert.True(summary.DeclaresGenerativeAi);
        Assert.Equal(new[] { "c2pa.created", "c2pa.converted" }, summary.Actions.Select(static action => action.Action));
        Assert.Equal("GPT-4o", summary.Actions[0].SoftwareAgent);
        Assert.Equal(OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia, summary.Actions[0].DigitalSourceKind);
        Assert.Equal("prompt-image.png", Assert.Single(summary.Ingredients));
    }

    [Fact]
    public void LegacyClaimGeneratorStringAndUnsignedManifestAreRead() {
        byte[] store = Store(Manifest("urn:uuid:only", claimGenerator: "make_test_images/0.16.1", actions: Actions(("c2pa.drawing", null, null)), signer: null));

        OfficeC2paManifestSummary? summary = OfficeC2paManifestStore.TryDescribe(store, 0, store.Length);

        Assert.NotNull(summary);
        Assert.Equal("make_test_images/0.16.1", summary!.ClaimGenerator);
        Assert.Null(summary.SignedBy);
        Assert.False(summary.DeclaresGenerativeAi);
        Assert.Equal("c2pa.drawing", Assert.Single(summary.Actions).Action);
    }

    [Fact]
    public void SummaryIsAttachedToEvidenceForAValidStore() {
        byte[] store = Store(Manifest("urn:uuid:only", claimGenerator: "ChatGPT", actions: Actions(("c2pa.created", "GPT-4o", TrainedAlgorithmicMedia)), signer: null));
        Assert.True(OfficeC2paManifestStore.IsValid(store, 0, store.Length, store.Length, 1024, out int length));
        Assert.Equal(store.Length, length);
        var evidence = new OfficeProvenanceEvidence(OfficeProvenanceCarrierKind.C2paManifest, "fixture", true, store.Length)
            .WithManifest(OfficeC2paManifestStore.TryDescribe(store, 0, store.Length));
        Assert.True(evidence.Manifest!.DeclaresGenerativeAi);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(17)]
    [InlineData(64)]
    public void TruncatedOrCorruptStoresNeverThrow(int keep) {
        byte[] store = Store(Manifest("urn:uuid:only", claimGenerator: "ChatGPT", actions: Actions(("c2pa.created", "GPT-4o", TrainedAlgorithmicMedia)),
            signer: Certificate("OpenAI", "Truepic")));
        byte[] truncated = store.Take(keep).ToArray();
        Assert.Null(OfficeC2paManifestStore.TryDescribe(truncated, 0, truncated.Length));

        byte[] corrupt = (byte[])store.Clone();
        for (int index = 40; index < corrupt.Length; index += 7) corrupt[index] ^= 0x5A;
        _ = OfficeC2paManifestStore.TryDescribe(corrupt, 0, corrupt.Length); // must not throw
    }

    [Fact]
    public void CborReaderRejectsOversizedLengthsAndDeepNesting() {
        // A text string claiming 4 GB, and 64 nested arrays.
        Assert.False(OfficeCborReader.TryDecode(new byte[] { 0x7A, 0xFF, 0xFF, 0xFF, 0xFF, 0x41 }, 0, 6, out _));
        byte[] nested = Enumerable.Repeat((byte)0x81, 64).Concat(new byte[] { 0x01 }).ToArray();
        Assert.False(OfficeCborReader.TryDecode(nested, 0, nested.Length, out _));
        Assert.True(OfficeCborReader.TryDecode(new byte[] { 0xA1, 0x61, 0x61, 0x01 }, 0, 4, out object? map));
        Assert.Equal(1L, ((Dictionary<object, object?>)map!)["a"]);
    }

    // ---- fixture builders: JUMBF boxes, CBOR, and a minimal DER certificate ----

    private static byte[] Store(params byte[][] manifests) =>
        Box("jumb", Join(Description("c2pa", "c2pa"), Join(manifests)));

    private static byte[] Manifest(
        string label,
        string? claimGenerator,
        byte[] actions,
        byte[]? signer,
        (string Name, string? Version)? claimGeneratorInfo = null,
        string? ingredient = null) {
        var claim = new List<(object, object?)> { ("dc:title", "image.png"), ("dc:format", "image/png") };
        if (claimGenerator != null) claim.Add(("claim_generator", claimGenerator));
        if (claimGeneratorInfo != null) claim.Add(("claim_generator_info", new object?[] { Map(("name", claimGeneratorInfo.Value.Name)) }));
        var assertions = new List<byte[]> { Assertion("c2pa.actions.v2", actions) };
        if (ingredient != null) assertions.Add(Assertion("c2pa.ingredient.v3", Cbor(Map(("dc:title", ingredient)))));
        object? x5chain = signer == null ? null : new object?[] { signer };
        byte[] protectedHeader = Cbor(x5chain == null ? Map((1L, -7L)) : Map((1L, -7L), (33L, x5chain)));
        byte[] sign1 = Cbor(new Tagged(18, new object?[] { protectedHeader, Map(), null, new byte[] { 1, 2, 3 } }));
        return Box("jumb", Join(
            Description("c2ma", label),
            Box("jumb", Join(Description("c2as", "c2pa.assertions"), Join(assertions.ToArray()))),
            Box("jumb", Join(Description("c2cl", "c2pa.claim.v2"), Box("cbor", Cbor(Map(claim.ToArray()))))),
            Box("jumb", Join(Description("c2cs", "c2pa.signature"), Box("cbor", sign1)))));
    }

    private static byte[] Actions(params (string Action, string? Agent, string? Source)[] actions) =>
        Cbor(Map(("actions", actions.Select(static action => {
            var entries = new List<(object, object?)> { ("action", action.Action) };
            if (action.Agent != null) entries.Add(("softwareAgent", Map(("name", action.Agent))));
            if (action.Source != null) entries.Add(("digitalSourceType", action.Source));
            return (object?)Map(entries.ToArray());
        }).ToArray())));

    private static byte[] Assertion(string label, byte[] cbor) => Box("jumb", Join(Description("cbor", label), Box("cbor", cbor)));

    private static byte[] Description(string code, string label) => Box("jumd", Join(
        Encoding.ASCII.GetBytes(code), new byte[] { 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 },
        new byte[] { 0x03 }, Encoding.UTF8.GetBytes(label + "\0")));

    private static byte[] Box(string type, byte[] payload) {
        int length = payload.Length + 8;
        return Join(new[] { (byte)(length >> 24), (byte)(length >> 16), (byte)(length >> 8), (byte)length }, Encoding.ASCII.GetBytes(type), payload);
    }

    private static byte[] Join(params byte[][] parts) => parts.SelectMany(static part => part).ToArray();

    private static Dictionary<object, object?> Map(params (object Key, object? Value)[] entries) {
        var map = new Dictionary<object, object?>();
        foreach ((object key, object? value) in entries) map[key] = value;
        return map;
    }

    private sealed record Tagged(ulong Tag, object? Value);

    private static byte[] Cbor(object? value) {
        var output = new List<byte>();
        Write(output, value);
        return output.ToArray();
    }

    private static void Write(List<byte> output, object? value) {
        switch (value) {
            case null: output.Add(0xF6); break;
            case long number when number >= 0: Head(output, 0, (ulong)number); break;
            case long number: Head(output, 1, (ulong)(-1 - number)); break;
            case string text: { byte[] bytes = Encoding.UTF8.GetBytes(text); Head(output, 3, (ulong)bytes.Length); output.AddRange(bytes); break; }
            case byte[] bytes: Head(output, 2, (ulong)bytes.Length); output.AddRange(bytes); break;
            case Tagged tagged: Head(output, 6, tagged.Tag); Write(output, tagged.Value); break;
            case Dictionary<object, object?> map:
                Head(output, 5, (ulong)map.Count);
                foreach (KeyValuePair<object, object?> entry in map) { Write(output, entry.Key); Write(output, entry.Value); }
                break;
            case object?[] items: Head(output, 4, (ulong)items.Length); foreach (object? item in items) Write(output, item); break;
            default: throw new ArgumentException("Unsupported test CBOR value.");
        }
    }

    private static void Head(List<byte> output, int major, ulong argument) {
        if (argument < 24) { output.Add((byte)(major << 5 | (int)argument)); return; }
        if (argument <= byte.MaxValue) { output.Add((byte)(major << 5 | 24)); output.Add((byte)argument); return; }
        output.Add((byte)(major << 5 | 25));
        output.Add((byte)(argument >> 8));
        output.Add((byte)argument);
    }

    /// <summary>A structurally shaped (unsigned) X.509 certificate carrying only the fields the reader walks.</summary>
    private static byte[] Certificate(string subjectOrganization, string issuerOrganization) {
        byte[] tbs = Der(0x30, Join(
            Der(0xA0, Der(0x02, new byte[] { 2 })),                       // version v3
            Der(0x02, new byte[] { 0x01 }),                               // serialNumber
            Der(0x30, Der(0x06, new byte[] { 0x2A, 0x86, 0x48, 0xCE, 0x3D, 0x04, 0x03, 0x02 })), // ecdsa-with-SHA256
            Name(issuerOrganization),
            Der(0x30, Join(Der(0x17, Encoding.ASCII.GetBytes("260101000000Z")), Der(0x17, Encoding.ASCII.GetBytes("270101000000Z")))),
            Name(subjectOrganization)));
        return Der(0x30, Join(tbs, Der(0x30, new byte[0]), Der(0x03, new byte[] { 0 })));
    }

    private static byte[] Name(string organization) => Der(0x30, Der(0x31, Der(0x30, Join(
        Der(0x06, new byte[] { 0x55, 0x04, 0x0A }), Der(0x0C, Encoding.UTF8.GetBytes(organization))))));

    private static byte[] Der(byte tag, byte[] content) {
        if (content.Length < 0x80) return Join(new[] { tag, (byte)content.Length }, content);
        return Join(new[] { tag, (byte)0x82, (byte)(content.Length >> 8), (byte)content.Length }, content);
    }
}

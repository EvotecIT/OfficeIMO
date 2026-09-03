using System.ComponentModel;
using System.IO;
using System.Text.Json;
using System.Threading;
using OfficeIMO.Provenance;

namespace OfficeIMO.Provenance.C2pa.Tests;

public sealed class C2paToolProvenanceSignerTests {
    [Fact]
    public void ProductionSigningUsesExternalSignerParentAndSafeManifest() {
        using var fixture = new SigningFixture();
        var runner = new SigningRunner(0, createEmbeddedManifest: true);
        var signer = new C2paToolProvenanceSigner("c2patool", "remote-signer", false, runner);
        var claim = new OfficeProvenanceClaim(
            "OfficeIMO/3.2.4",
            new[] {
                new OfficeProvenanceAction(OfficeProvenanceActionKind.Opened),
                new OfficeProvenanceAction(
                    OfficeProvenanceActionKind.Edited,
                    OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia),
                new OfficeProvenanceAction(OfficeProvenanceActionKind.Published)
            },
            "Derived asset");

        OfficeProvenanceSigningResult result = signer.Sign(
            new OfficeProvenanceSigningRequest(fixture.Input, fixture.Output, claim, fixture.Parent),
            new OfficeProvenanceSigningOptions { IncludeRawReport = true });

        Assert.Equal(OfficeProvenanceSigningStatus.Signed, result.Status);
        Assert.True(result.StructuralReport!.HasC2paManifest);
        Assert.Equal(fixture.Output, result.OutputPath);
        Assert.Equal("signed", File.ReadAllText(fixture.Output).Split('\n')[1]);
        Assert.Equal("tool output", result.RawReport);
        Assert.NotNull(runner.Request);
        Assert.Equal(new[] { "--version" }, runner.VersionRequest!.Arguments);
        Assert.Equal("remote-signer", ValueAfter(runner.Request!.Arguments, "--signer-path"));
        Assert.Equal(fixture.Parent, ValueAfter(runner.Request.Arguments, "--parent"));
        Assert.DoesNotContain("--create", runner.Request.Arguments);
        Assert.DoesNotContain("private_key", runner.ManifestJson, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("certificate", runner.ManifestJson, StringComparison.OrdinalIgnoreCase);
        using JsonDocument manifest = JsonDocument.Parse(runner.ManifestJson);
        JsonElement root = manifest.RootElement;
        Assert.Equal(
            "OfficeIMO/3.2.4",
            root.GetProperty("claim_generator_info")[0].GetProperty("name").GetString());
        Assert.Equal("Derived asset", root.GetProperty("title").GetString());
        JsonElement assertion = root.GetProperty("assertions")[0];
        Assert.Equal("c2pa.actions.v2", assertion.GetProperty("label").GetString());
        JsonElement actions = assertion.GetProperty("data").GetProperty("actions");
        Assert.Equal("c2pa.edited", actions[0].GetProperty("action").GetString());
        Assert.Equal(
            "http://cv.iptc.org/newscodes/digitalsourcetype/compositeWithTrainedAlgorithmicMedia",
            actions[0].GetProperty("digitalSourceType").GetString());
        Assert.Equal("c2pa.published", actions[1].GetProperty("action").GetString());
        Assert.False(File.Exists(runner.ManifestPath));
    }

    [Fact]
    public void DevelopmentCredentialsAreExplicitAndReported() {
        using var fixture = new SigningFixture();
        var runner = new SigningRunner(0, createEmbeddedManifest: true);
        var signer = new C2paToolProvenanceSigner("c2patool", null, true, runner);

        OfficeProvenanceSigningResult result = signer.Sign(fixture.Request());

        Assert.True(result.Succeeded);
        Assert.True(signer.UsesBuiltInTestCredentials);
        Assert.Null(signer.SignerPath);
        Assert.DoesNotContain("--signer-path", runner.Request!.Arguments);
        Assert.Equal("digitalCapture", ValueAfter(runner.Request.Arguments, "--create"));
        using JsonDocument manifest = JsonDocument.Parse(runner.ManifestJson);
        Assert.Equal(0, manifest.RootElement.GetProperty("assertions").GetArrayLength());
        Assert.Contains(result.Findings, finding => finding.Contains("development credentials", StringComparison.Ordinal));
    }

    [Fact]
    public void UnsupportedProviderVersionIsUnavailableBeforeSigning() {
        using var fixture = new SigningFixture();
        var runner = new SigningRunner(0, createEmbeddedManifest: true, version: "c2patool 0.26.9");
        var signer = new C2paToolProvenanceSigner("c2patool", "remote-signer", false, runner);

        OfficeProvenanceSigningResult result = signer.Sign(fixture.Request());

        Assert.Equal(OfficeProvenanceSigningStatus.ProviderUnavailable, result.Status);
        Assert.Contains(result.Findings, finding => finding.Contains("0.27.0", StringComparison.Ordinal));
        Assert.Null(runner.Request);
        Assert.False(File.Exists(fixture.Output));
    }

    [Fact]
    public void ProviderFailureDeletesPartialOutputAndPreservesDestination() {
        using var fixture = new SigningFixture();
        File.WriteAllText(fixture.Output, "existing");
        var runner = new SigningRunner(2, createEmbeddedManifest: false, createPartialOutput: true);
        var signer = new C2paToolProvenanceSigner("c2patool", "remote-signer", false, runner);

        OfficeProvenanceSigningResult result = signer.Sign(fixture.Request());

        Assert.Equal(OfficeProvenanceSigningStatus.Rejected, result.Status);
        Assert.Equal("existing", File.ReadAllText(fixture.Output));
        Assert.False(File.Exists(runner.StagingPath));
    }

    [Fact]
    public void SuccessfulExitWithoutEmbeddedManifestIsRejectedAndNotCommitted() {
        using var fixture = new SigningFixture();
        var runner = new SigningRunner(0, createEmbeddedManifest: false, createPartialOutput: true);
        var signer = new C2paToolProvenanceSigner("c2patool", "remote-signer", false, runner);

        OfficeProvenanceSigningResult result = signer.Sign(fixture.Request());

        Assert.Equal(OfficeProvenanceSigningStatus.Error, result.Status);
        Assert.False(File.Exists(fixture.Output));
        Assert.False(File.Exists(runner.StagingPath));
    }

    [Fact]
    public void SuccessfulExitWithMalformedEmbeddedManifestIsRejectedAndNotCommitted() {
        using var fixture = new SigningFixture();
        var runner = new SigningRunner(0, createEmbeddedManifest: false, createMalformedManifest: true);
        var signer = new C2paToolProvenanceSigner("c2patool", "remote-signer", false, runner);

        OfficeProvenanceSigningResult result = signer.Sign(fixture.Request());

        Assert.Equal(OfficeProvenanceSigningStatus.Error, result.Status);
        Assert.Contains(result.Findings, finding => finding.Contains("structurally valid", StringComparison.Ordinal));
        Assert.False(File.Exists(fixture.Output));
        Assert.False(File.Exists(runner.StagingPath));
    }

    [Fact]
    public void MissingProviderIsNormalizedWithoutCreatingOutput() {
        using var fixture = new SigningFixture();
        var signer = new C2paToolProvenanceSigner(
            "missing-c2patool",
            "remote-signer",
            false,
            new UnavailableRunner());

        OfficeProvenanceSigningResult result = signer.Sign(fixture.Request());

        Assert.Equal(OfficeProvenanceSigningStatus.ProviderUnavailable, result.Status);
        Assert.False(File.Exists(fixture.Output));
    }

    [Fact]
    public void SigningRejectsInPlaceAndFormatChangingOutputs() {
        using var fixture = new SigningFixture();
        var signer = new C2paToolProvenanceSigner(
            "c2patool",
            "remote-signer",
            false,
            new SigningRunner(0, createEmbeddedManifest: true));

        Assert.Throws<ArgumentException>(() => signer.Sign(new OfficeProvenanceSigningRequest(
            fixture.Input,
            fixture.Input,
            fixture.Claim)));
        Assert.Throws<ArgumentException>(() => signer.Sign(new OfficeProvenanceSigningRequest(
            fixture.Input,
            Path.ChangeExtension(fixture.Output, ".png"),
            fixture.Claim)));
    }

    [Fact]
    public void ClaimActionsRejectUndefinedKinds() {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new OfficeProvenanceAction((OfficeProvenanceActionKind)int.MaxValue));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new OfficeProvenanceAction(
                OfficeProvenanceActionKind.Created,
                (OfficeProvenanceDigitalSourceKind)int.MaxValue));
    }

    [Fact]
    public void SigningRejectsClaimsWhoseFirstActionDoesNotMatchIntent() {
        using var fixture = new SigningFixture();
        var signer = new C2paToolProvenanceSigner(
            "c2patool",
            "remote-signer",
            false,
            new SigningRunner(0, createEmbeddedManifest: true));
        var createdWithoutSource = new OfficeProvenanceClaim(
            "OfficeIMO/Tests",
            new[] { new OfficeProvenanceAction(OfficeProvenanceActionKind.Created) });
        var derivedWithoutOpened = new OfficeProvenanceClaim(
            "OfficeIMO/Tests",
            new[] { new OfficeProvenanceAction(OfficeProvenanceActionKind.Edited) });

        Assert.Throws<ArgumentException>(() => signer.Sign(
            new OfficeProvenanceSigningRequest(fixture.Input, fixture.Output, createdWithoutSource)));
        Assert.Throws<ArgumentException>(() => signer.Sign(
            new OfficeProvenanceSigningRequest(fixture.Input, fixture.Output, derivedWithoutOpened, fixture.Parent)));
    }

    private static string ValueAfter(IReadOnlyList<string> arguments, string name) {
        int index = arguments.ToList().IndexOf(name);
        Assert.InRange(index, 0, arguments.Count - 2);
        return arguments[index + 1];
    }

    private sealed class SigningRunner : IC2paToolProcessRunner {
        private readonly int _exitCode;
        private readonly bool _createEmbeddedManifest;
        private readonly bool _createPartialOutput;
        private readonly bool _createMalformedManifest;
        private readonly string _version;

        internal SigningRunner(
            int exitCode,
            bool createEmbeddedManifest,
            bool createPartialOutput = false,
            bool createMalformedManifest = false,
            string version = "c2patool 0.27.15") {
            _exitCode = exitCode;
            _createEmbeddedManifest = createEmbeddedManifest;
            _createPartialOutput = createPartialOutput;
            _createMalformedManifest = createMalformedManifest;
            _version = version;
        }

        internal C2paToolProcessRequest? Request { get; private set; }
        internal C2paToolProcessRequest? VersionRequest { get; private set; }
        internal string ManifestJson { get; private set; } = string.Empty;
        internal string ManifestPath { get; private set; } = string.Empty;
        internal string StagingPath { get; private set; } = string.Empty;

        public C2paToolProcessResult Run(C2paToolProcessRequest request, CancellationToken cancellationToken = default) {
            if (request.Arguments.SequenceEqual(new[] { "--version" })) {
                VersionRequest = request;
                return new C2paToolProcessResult(0, _version, string.Empty);
            }
            Request = request;
            ManifestPath = ValueAfter(request.Arguments, "--manifest");
            StagingPath = ValueAfter(request.Arguments, "--output");
            ManifestJson = File.ReadAllText(ManifestPath);
            if (_createEmbeddedManifest) {
                string encoded = Convert.ToBase64String(CreateManifestStore());
                File.WriteAllText(
                    StagingPath,
                    "# -----BEGIN C2PA MANIFEST----- data:application/c2pa;base64," + encoded +
                    " -----END C2PA MANIFEST-----\nsigned");
            } else if (_createMalformedManifest) {
                File.WriteAllText(
                    StagingPath,
                    "# -----BEGIN C2PA MANIFEST----- data:application/c2pa;base64,not-base64 -----END C2PA MANIFEST-----\n");
            } else if (_createPartialOutput) {
                File.WriteAllText(StagingPath, "partial");
            }
            return new C2paToolProcessResult(_exitCode, "tool output", _exitCode == 0 ? string.Empty : "tool error");
        }
    }

    private sealed class UnavailableRunner : IC2paToolProcessRunner {
        public C2paToolProcessResult Run(C2paToolProcessRequest request, CancellationToken cancellationToken = default) =>
            throw new Win32Exception("Executable not found.");
    }

    private sealed class SigningFixture : IDisposable {
        internal SigningFixture() {
            DirectoryPath = Path.Combine(Path.GetTempPath(), "OfficeIMO-Signing-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(DirectoryPath);
            Input = Path.Combine(DirectoryPath, "input.txt");
            Parent = Path.Combine(DirectoryPath, "parent.txt");
            Output = Path.Combine(DirectoryPath, "output.txt");
            File.WriteAllText(Input, "input");
            File.WriteAllText(Parent, "parent");
            Claim = new OfficeProvenanceClaim(
                "OfficeIMO/3.2.4",
                new[] {
                    new OfficeProvenanceAction(
                        OfficeProvenanceActionKind.Created,
                        OfficeProvenanceDigitalSourceKind.DigitalCapture)
                });
        }

        internal string DirectoryPath { get; }
        internal string Input { get; }
        internal string Parent { get; }
        internal string Output { get; }
        internal OfficeProvenanceClaim Claim { get; }
        internal OfficeProvenanceSigningRequest Request() => new(Input, Output, Claim);

        public void Dispose() {
            try { Directory.Delete(DirectoryPath, recursive: true); } catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static byte[] CreateManifestStore() {
        byte[] storeDescription = CreateBox("jumd", Join(
            C2paUuid("c2pa"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa\0")));
        byte[] manifestDescription = CreateBox("jumd", Join(
            C2paUuid("c2ma"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("m\0")));
        byte[] assertionStoreDescription = CreateBox("jumd", Join(
            C2paUuid("c2as"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.assertions\0")));
        byte[] assertionDescription = CreateBox("jumd", Join(
            C2paUuid("c2ac"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.test\0")));
        byte[] assertionStore = CreateBox("jumb", Join(assertionStoreDescription,
            CreateBox("jumb", Join(assertionDescription, CreateBox("cbor", new byte[] { 0xA0 })))));
        byte[] claimDescription = CreateBox("jumd", Join(
            C2paUuid("c2cl"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.claim\0")));
        byte[] claim = CreateBox("jumb", Join(claimDescription, CreateBox("cbor", new byte[] { 0xA0 })));
        byte[] signatureDescription = CreateBox("jumd", Join(
            C2paUuid("c2cs"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.signature\0")));
        byte[] signature = CreateBox("jumb", Join(signatureDescription, CreateBox("cbor", new byte[] { 0xA0 })));
        return CreateBox("jumb", Join(storeDescription,
            CreateBox("jumb", Join(manifestDescription, assertionStore, claim, signature))));
    }

    private static byte[] C2paUuid(string code) => Join(
        Encoding.ASCII.GetBytes(code),
        new byte[] { 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 });

    private static byte[] CreateBox(string type, byte[] payload) {
        byte[] box = new byte[payload.Length + 8];
        WriteBigEndian(box, 0, box.Length);
        Encoding.ASCII.GetBytes(type).CopyTo(box, 4);
        Buffer.BlockCopy(payload, 0, box, 8, payload.Length);
        return box;
    }

    private static byte[] Join(params byte[][] arrays) {
        byte[] result = new byte[arrays.Sum(array => array.Length)];
        int offset = 0;
        foreach (byte[] array in arrays) {
            Buffer.BlockCopy(array, 0, result, offset, array.Length);
            offset += array.Length;
        }
        return result;
    }

    private static void WriteBigEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}

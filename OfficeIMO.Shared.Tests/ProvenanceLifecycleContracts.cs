using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ProvenanceLifecycleContracts {
    [Fact]
    public void DefaultPolicyBlocksChangedCredentialedContent() {
        string source = CreateTextAsset("body\n");
        string candidate = CreateTextAsset("changed\n");
        string output = TemporaryPath();
        try {
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                OfficeProvenanceLifecycle.FinalizeFile(source, candidate, output));

            Assert.Contains("Content Credential", exception.Message, StringComparison.Ordinal);
            Assert.False(File.Exists(output));
        } finally {
            Delete(source, candidate, output);
        }
    }

    [Fact]
    public void RemovalPolicyStripsCarriedCredentialAndRetainsAuditEvidence() {
        string source = CreateTextAsset("body\n");
        string candidate = CreateTextAsset("changed\n");
        string output = TemporaryPath();
        try {
            var options = new OfficeProvenanceTransformationOptions {
                Policy = OfficeProvenanceTransformationPolicy.RemoveInvalidated
            };

            OfficeProvenanceTransformationResult result = OfficeProvenanceLifecycle.FinalizeFile(
                source,
                candidate,
                output,
                options);

            Assert.True(result.ContentChanged);
            Assert.True(result.Source.HasC2paManifest);
            Assert.True(result.Candidate.HasC2paManifest);
            Assert.NotNull(result.Removal);
            Assert.True(result.Removal!.WasChanged);
            Assert.False(result.Output.HasC2paManifest);
            Assert.Equal("changed\n", File.ReadAllText(output, Encoding.UTF8));
        } finally {
            Delete(source, candidate, output);
        }
    }

    [Fact]
    public void DerivedSigningAlwaysPassesTheSourceAsParentIngredient() {
        string source = CreateTextAsset("body\n");
        string candidate = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(candidate, "changed\n", new UTF8Encoding(false));
        string output = TemporaryPath();
        string expectedSource = File.ReadAllText(source, Encoding.UTF8);
        string expectedCandidate = File.ReadAllText(candidate, Encoding.UTF8);
        try {
            var options = new OfficeProvenanceTransformationOptions {
                Policy = OfficeProvenanceTransformationPolicy.SignAsDerived,
                Claim = new OfficeProvenanceClaim(
                    "OfficeIMO/Tests",
                    new[] {
                        new OfficeProvenanceAction(OfficeProvenanceActionKind.Opened),
                        new OfficeProvenanceAction(OfficeProvenanceActionKind.Edited)
                    })
            };
            var signer = new StubSigner(source, candidate);

            OfficeProvenanceTransformationResult result = OfficeProvenanceLifecycle.FinalizeFile(
                source,
                candidate,
                output,
                options,
                signer);

            Assert.NotEqual(Path.GetFullPath(source), Path.GetFullPath(signer.Request!.ParentPath!));
            Assert.NotEqual(Path.GetFullPath(candidate), Path.GetFullPath(signer.Request.InputPath));
            Assert.NotEqual(Path.GetFullPath(output), Path.GetFullPath(signer.Request.OutputPath));
            Assert.Equal(Path.GetFileName(source), Path.GetFileName(signer.Request.ParentPath));
            Assert.Equal(Path.GetFileName(candidate), Path.GetFileName(signer.Request.InputPath));
            Assert.Equal(expectedSource, signer.ParentText);
            Assert.Equal(expectedCandidate, signer.InputText);
            Assert.Equal("mutated source", File.ReadAllText(source));
            Assert.Equal("mutated candidate", File.ReadAllText(candidate));
            Assert.False(File.Exists(signer.Request.ParentPath));
            Assert.False(File.Exists(signer.Request.InputPath));
            Assert.False(File.Exists(signer.Request.OutputPath));
            Assert.True(result.Signing!.Succeeded);
            Assert.Equal(Path.GetFullPath(output), Path.GetFullPath(result.Signing.OutputPath!));
            Assert.True(result.Output.HasExternalC2paManifest);
        } finally {
            Delete(source, candidate, output);
        }
    }

    [Fact]
    public void ChangedCandidateCannotOverwriteItsProvenanceSource() {
        string source = CreateTextAsset("body\n");
        string candidate = CreateTextAsset("changed\n");
        try {
            var options = new OfficeProvenanceTransformationOptions {
                Policy = OfficeProvenanceTransformationPolicy.RemoveInvalidated
            };

            Assert.Throws<InvalidOperationException>(() =>
                OfficeProvenanceLifecycle.FinalizeFile(source, candidate, source, options));
        } finally {
            Delete(source, candidate);
        }
    }

    [Fact]
    public void ChangedUncredentialedContentCanUseTheDefaultPolicy() {
        string source = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        string candidate = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        string output = TemporaryPath();
        File.WriteAllText(source, "before", new UTF8Encoding(false));
        File.WriteAllText(candidate, "after", new UTF8Encoding(false));
        try {
            OfficeProvenanceTransformationResult result = OfficeProvenanceLifecycle.FinalizeFile(source, candidate, output);

            Assert.True(result.ContentChanged);
            Assert.Equal("after", File.ReadAllText(output, Encoding.UTF8));
        } finally {
            Delete(source, candidate, output);
        }
    }

    [Theory]
    [InlineData("provider")]
    [InlineData("path")]
    [InlineData("missing")]
    [InlineData("evidence")]
    [InlineData("malformed")]
    public void DerivedSigningValidatesActualProviderOutputBeforeCommit(string defect) {
        string source = CreateTextAsset("body\n");
        string candidate = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        string output = TemporaryPath();
        File.WriteAllText(candidate, "changed\n", new UTF8Encoding(false));
        File.WriteAllText(output, "existing", new UTF8Encoding(false));
        try {
            var options = new OfficeProvenanceTransformationOptions {
                Policy = OfficeProvenanceTransformationPolicy.SignAsDerived,
                Claim = new OfficeProvenanceClaim(
                    "OfficeIMO/Tests",
                    new[] { new OfficeProvenanceAction(OfficeProvenanceActionKind.Opened) })
            };

            Assert.Throws<InvalidOperationException>(() => OfficeProvenanceLifecycle.FinalizeFile(
                source,
                candidate,
                output,
                options,
                new DefectiveSigner(defect)));

            Assert.Equal("existing", File.ReadAllText(output));
        } finally {
            Delete(source, candidate, output);
        }
    }

    private static string CreateTextAsset(string body) {
        string path = TemporaryPath();
        string manifest = Convert.ToBase64String(ProvenanceCoreContracts.CreateManifestStoreForLifecycleTests());
        File.WriteAllText(
            path,
            "-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + manifest + "\n" +
            "-----END C2PA MANIFEST-----\n" + body,
            new UTF8Encoding(false));
        return path;
    }

    private static string TemporaryPath() =>
        Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");

    private static void Delete(params string[] paths) {
        foreach (string path in paths) if (File.Exists(path)) File.Delete(path);
    }

    private sealed class StubSigner : IOfficeProvenanceSigner {
        private readonly string _originalSource;
        private readonly string _originalCandidate;

        internal StubSigner(string originalSource, string originalCandidate) {
            _originalSource = originalSource;
            _originalCandidate = originalCandidate;
        }

        public string Name => "stub-signer";
        internal OfficeProvenanceSigningRequest? Request { get; private set; }
        internal string InputText { get; private set; } = string.Empty;
        internal string ParentText { get; private set; } = string.Empty;

        public OfficeProvenanceSigningResult Sign(
            OfficeProvenanceSigningRequest request,
            OfficeProvenanceSigningOptions? options = null) {
            Request = request;
            InputText = File.ReadAllText(request.InputPath, Encoding.UTF8);
            ParentText = File.ReadAllText(request.ParentPath!, Encoding.UTF8);
            File.WriteAllText(_originalSource, "mutated source", new UTF8Encoding(false));
            File.WriteAllText(_originalCandidate, "mutated candidate", new UTF8Encoding(false));
            byte[] output = Encoding.UTF8.GetBytes(
                "# -----BEGIN C2PA MANIFEST----- https://example.test/derived.c2pa -----END C2PA MANIFEST-----\n" +
                InputText);
            File.WriteAllBytes(request.OutputPath, output);
            OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(output, request.OutputPath);
            return new OfficeProvenanceSigningResult(
                OfficeProvenanceSigningStatus.Signed,
                Name,
                Array.Empty<string>(),
                request.OutputPath,
                report);
        }
    }

    private sealed class DefectiveSigner : IOfficeProvenanceSigner {
        private readonly string _defect;

        internal DefectiveSigner(string defect) => _defect = defect;

        public string Name => "defective-signer";

        public OfficeProvenanceSigningResult Sign(
            OfficeProvenanceSigningRequest request,
            OfficeProvenanceSigningOptions? options = null) {
            byte[] credentialed = Encoding.UTF8.GetBytes(
                "# -----BEGIN C2PA MANIFEST----- https://example.test/derived.c2pa -----END C2PA MANIFEST-----\nbody\n");
            if (_defect != "missing") {
                byte[] output = _defect switch {
                    "evidence" => Encoding.UTF8.GetBytes("unsigned"),
                    "malformed" => Encoding.UTF8.GetBytes(
                        "# -----BEGIN C2PA MANIFEST----- data:application/c2pa;base64,not-base64 -----END C2PA MANIFEST-----\n"),
                    _ => credentialed
                };
                File.WriteAllBytes(request.OutputPath, output);
            }
            OfficeProvenanceReport claimedReport = OfficeProvenanceInspector.Inspect(credentialed, request.OutputPath);
            return new OfficeProvenanceSigningResult(
                OfficeProvenanceSigningStatus.Signed,
                _defect == "provider" ? "some-other-provider" : Name,
                Array.Empty<string>(),
                _defect == "path" ? request.OutputPath + ".other" : request.OutputPath,
                claimedReport);
        }
    }
}

using System.ComponentModel;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using OfficeIMO.Provenance;

namespace OfficeIMO.Security.Tests;

public sealed class C2paToolProvenanceVerifierTests {
    [Theory]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\"}", 0, OfficeProvenanceVerificationStatus.Valid)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"code\":\"claimSignature.validated\"},{\"code\":\"assertion.dataHash.match\"},{\"code\":\"signingCredential.trusted\"}]}", 0, OfficeProvenanceVerificationStatus.Valid)]
    [InlineData("{\"active_manifest\":null,\"manifests\":{}}", 0, OfficeProvenanceVerificationStatus.NotPresent)]
    [InlineData("{\"active_manifest\":\"\"}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":\"   \"}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("[]", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":{}}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"code\":\"signingCredential.untrusted\"}]}", 0, OfficeProvenanceVerificationStatus.Untrusted)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"code\":\"assertion.dataHash.mismatch\"}]}", 0, OfficeProvenanceVerificationStatus.Invalid)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"code\":\"assertion.dataHash.mismatch\",\"success\":true}]}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"code\":\"assertion.dataHash.match\",\"success\":false}]}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":null,\"validation_status\":[{\"code\":\"assertion.dataHash.mismatch\"}]}", 0, OfficeProvenanceVerificationStatus.Invalid)]
    [InlineData("{\"active_manifest\":null,\"validation_status\":[{\"code\":\"signingCredential.untrusted\"}]}", 0, OfficeProvenanceVerificationStatus.Untrusted)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"code\":\"signingCredential.ocsp.revoked\"}]}", 0, OfficeProvenanceVerificationStatus.Untrusted)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"code\":\"signingCredential.expired\"}]}", 0, OfficeProvenanceVerificationStatus.Untrusted)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"code\":\"ingredient.manifest.missing\"}]}", 0, OfficeProvenanceVerificationStatus.Invalid)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":{}}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":\"valid\"}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[true]}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"code\":\"claimSignature.validated\",\"success\":\"true\"}]}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"code\":\"claimSignature.validated\",\"success\":null}]}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"code\":1,\"success\":true}]}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"success\":true}]}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"code\":\"   \",\"success\":true}]}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"code\":\"assertion.dataHash.mismatch\"}],\"validation_status\":[]}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":null,\"active_manifest\":\"urn:c2pa:test\"}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{\"success\":false,\"success\":true}]}", 0, OfficeProvenanceVerificationStatus.Error)]
    [InlineData("{\"active_manifest\":null,\"manifests\":{\"urn:c2pa:ingredient\":{\"active_manifest\":\"nested\"}}}", 0, OfficeProvenanceVerificationStatus.NotPresent)]
    [InlineData("not-json", 1, OfficeProvenanceVerificationStatus.Error)]
    public void VerifyNormalizesBoundedToolReports(string report, int exitCode, OfficeProvenanceVerificationStatus expected) {
        string assetPath = CreateAsset();
        try {
            var runner = new StubRunner(new C2paToolProcessResult(exitCode, report, "tool error"));
            var verifier = new C2paToolProvenanceVerifier("c2patool", runner);

            OfficeProvenanceVerificationResult result = verifier.Verify(assetPath, new OfficeProvenanceVerificationOptions {
                IncludeRawReport = true
            });

            Assert.Equal(expected, result.Status);
            Assert.Equal("c2patool", result.ProviderName);
            Assert.Equal(report, result.RawReport);
            Assert.NotNull(runner.Request);
            Assert.Equal(assetPath, runner.Request!.Arguments[0]);
            Assert.Equal("--settings", runner.Request.Arguments[1]);
            Assert.Contains("\"remote_manifest_fetch\":false", runner.SettingsJson, StringComparison.Ordinal);
            Assert.Contains("\"ocsp_fetch\":false", runner.SettingsJson, StringComparison.Ordinal);
        } finally {
            File.Delete(assetPath);
        }
    }

    [Fact]
    public void VerifyKeepsRawReportPrivateByDefault() {
        string assetPath = CreateAsset();
        try {
            var runner = new StubRunner(new C2paToolProcessResult(0, "{\"active_manifest\":\"urn:c2pa:test\"}", string.Empty));
            var verifier = new C2paToolProvenanceVerifier("c2patool", runner);

            OfficeProvenanceVerificationResult result = verifier.Verify(assetPath);

            Assert.Null(result.RawReport);
        } finally {
            File.Delete(assetPath);
        }
    }

    [Fact]
    public void VerifyIgnoresValidationStatusInsideCustomAssertionPayloads() {
        string assetPath = CreateAsset();
        try {
            const string report = "{\"active_manifest\":\"urn:c2pa:test\",\"manifests\":{\"urn:c2pa:test\":{\"assertions\":[{\"data\":{\"validation_status\":[{\"code\":\"signingCredential.untrusted\"}]}}]}}}";
            var verifier = new C2paToolProvenanceVerifier(
                "c2patool",
                new StubRunner(new C2paToolProcessResult(0, report, string.Empty)));

            OfficeProvenanceVerificationResult result = verifier.Verify(assetPath);

            Assert.Equal(OfficeProvenanceVerificationStatus.Valid, result.Status);
            Assert.Empty(result.Findings);
        } finally {
            File.Delete(assetPath);
        }
    }

    [Fact]
    public void VerifyDeduplicatesManyFindingsWhilePreservingFirstSeenOrder() {
        string assetPath = CreateAsset();
        try {
            string statuses = string.Join(",", Enumerable.Range(0, 5000)
                .Select(index => $"{{\"code\":\"failure.{index}\"}}"));
            string report = $"{{\"active_manifest\":\"urn:c2pa:test\",\"validation_status\":[{statuses},{{\"code\":\"failure.0\"}}]}}";
            var verifier = new C2paToolProvenanceVerifier(
                "c2patool",
                new StubRunner(new C2paToolProcessResult(0, report, string.Empty)));

            OfficeProvenanceVerificationResult result = verifier.Verify(assetPath);

            Assert.Equal(5000, result.Findings.Count);
            Assert.Equal("failure.0", result.Findings[0]);
            Assert.Equal("failure.4999", result.Findings[4999]);
        } finally {
            File.Delete(assetPath);
        }
    }

    [Fact]
    public void VerifyAddsLocalTrustArgumentsWithoutEnablingNetwork() {
        string assetPath = CreateAsset();
        string anchorsPath = CreateAsset();
        try {
            var runner = new StubRunner(new C2paToolProcessResult(0, "{\"active_manifest\":\"urn:c2pa:test\"}", string.Empty));
            var verifier = new C2paToolProvenanceVerifier("c2patool", runner);

            verifier.Verify(assetPath, new OfficeProvenanceVerificationOptions { TrustAnchorsPath = anchorsPath });

            Assert.NotNull(runner.Request);
            Assert.Equal("trust", runner.Request!.Arguments[3]);
            Assert.Equal("--trust_anchors", runner.Request.Arguments[4]);
            Assert.Equal(Path.GetFileName(anchorsPath), runner.Request.Arguments[5]);
            Assert.Contains("\"remote_manifest_fetch\":false", runner.SettingsJson, StringComparison.Ordinal);
        } finally {
            File.Delete(assetPath);
            File.Delete(anchorsPath);
        }
    }

    [Fact]
    public void CrossVolumeTrustPathsRemainAbsoluteOnWindows() {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;

        string path = C2paToolProvenanceVerifier.GetRelativePath(@"C:\work\verification", @"D:\trust\anchors.pem");

        Assert.Equal(@"D:\trust\anchors.pem", path, ignoreCase: true);
    }

    [Fact]
    public void VerifyRejectsRemoteTrustMaterialUnlessNetworkWasExplicitlyEnabled() {
        string assetPath = CreateAsset();
        try {
            var verifier = new C2paToolProvenanceVerifier("c2patool", new StubRunner(
                new C2paToolProcessResult(0, "{}", string.Empty)));

            Assert.Throws<ArgumentException>(() => verifier.Verify(assetPath, new OfficeProvenanceVerificationOptions {
                TrustAnchorsPath = "https://example.test/anchors.pem"
            }));
        } finally {
            File.Delete(assetPath);
        }
    }

    [Fact]
    public void VerifyReportsAnUnavailableExecutableWithoutThrowing() {
        string assetPath = CreateAsset();
        try {
            var verifier = new C2paToolProvenanceVerifier("missing-c2patool", new ThrowingRunner());

            OfficeProvenanceVerificationResult result = verifier.Verify(assetPath);

            Assert.Equal(OfficeProvenanceVerificationStatus.ProviderUnavailable, result.Status);
            Assert.Single(result.Findings);
        } finally {
            File.Delete(assetPath);
        }
    }

    [Fact]
    public void DefaultRunnerReportsAnUnavailableExplicitExecutable() {
        string assetPath = CreateAsset();
        string executablePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "c2patool");
        try {
            var verifier = new C2paToolProvenanceVerifier(executablePath);

            OfficeProvenanceVerificationResult result = verifier.Verify(assetPath);

            Assert.Equal(OfficeProvenanceVerificationStatus.ProviderUnavailable, result.Status);
            Assert.Single(result.Findings);
        } finally {
            File.Delete(assetPath);
        }
    }

    [Fact]
    public void ProcessRunnerBoundsInheritedOutputHandlesAfterTheParentExits() {
        string executable;
        IReadOnlyList<string> arguments;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            executable = Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe";
            arguments = new[] { "/d", "/c", "start /b ping 127.0.0.1 -n 6" };
        } else {
            executable = "/bin/sh";
            arguments = new[] { "-c", "(sleep 5) & exit 0" };
        }
        var request = new C2paToolProcessRequest(
            executable,
            arguments,
            Path.GetTempPath(),
            TimeSpan.FromMilliseconds(300),
            1024 * 1024);
        var timer = Stopwatch.StartNew();

        Assert.Throws<TimeoutException>(() => new C2paToolProcessRunner().Run(request));

        Assert.True(timer.Elapsed < TimeSpan.FromSeconds(3), $"Runner blocked for {timer.Elapsed}.");
    }

    [Fact]
    public void ProcessRunnerTearsDownInheritedOutputHandlesWhenTheParentTimesOut() {
        string executable;
        IReadOnlyList<string> arguments;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            executable = Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe";
            arguments = new[] { "/d", "/c", "start /b ping 127.0.0.1 -n 6 & ping 127.0.0.1 -n 6" };
        } else {
            executable = "/bin/sh";
            arguments = new[] { "-c", "(sleep 5) & sleep 5" };
        }
        var request = new C2paToolProcessRequest(
            executable,
            arguments,
            Path.GetTempPath(),
            TimeSpan.FromMilliseconds(300),
            1024 * 1024);
        var timer = Stopwatch.StartNew();

        Assert.Throws<TimeoutException>(() => new C2paToolProcessRunner().Run(request));

        Assert.True(timer.Elapsed < TimeSpan.FromSeconds(3), $"Runner blocked for {timer.Elapsed}.");
    }

    [Fact]
    public void UnixShellContainmentKillsOrphanedChildrenWithoutSetsid() {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        string marker = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".orphan");
        try {
            var request = new C2paToolProcessRequest(
                "/bin/sh",
                new[] { "-c", "(sleep 1; printf orphan > \"$1\") & exit 0", "officeimo-test", marker },
                Path.GetTempPath(),
                TimeSpan.FromSeconds(2),
                1024 * 1024);

            C2paToolProcessResult result = new C2paToolProcessRunner(useExternalUnixSessionLauncher: false).Run(request);
            Thread.Sleep(1500);

            Assert.Equal(0, result.ExitCode);
            Assert.False(File.Exists(marker));
        } finally {
            if (File.Exists(marker)) File.Delete(marker);
        }
    }

    private static string CreateAsset() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".jpg");
        File.WriteAllBytes(path, new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 });
        return path;
    }

    private sealed class StubRunner : IC2paToolProcessRunner {
        private readonly C2paToolProcessResult _result;
        internal StubRunner(C2paToolProcessResult result) => _result = result;
        internal C2paToolProcessRequest? Request { get; private set; }
        internal string SettingsJson { get; private set; } = string.Empty;
        public C2paToolProcessResult Run(C2paToolProcessRequest request) {
            Request = request;
            SettingsJson = File.ReadAllText(request.Arguments[2]);
            return _result;
        }
    }

    private sealed class ThrowingRunner : IC2paToolProcessRunner {
        public C2paToolProcessResult Run(C2paToolProcessRequest request) =>
            throw new Win32Exception("Executable not found.");
    }
}

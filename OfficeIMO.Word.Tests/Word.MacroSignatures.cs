using System.IO.Packaging;
using System.Security.Cryptography.X509Certificates;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Security;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData("docm")]
        [InlineData("dotm")]
        public void MacroSignatureInspectorReportsUnsignedProjectCrossPlatform(string extension) {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureUnsigned." + extension);

            WordMacroProjectSignatureInfo info = WordDocument.InspectMacroProjectSignatures(filePath);

            Assert.True(info.IsMacroEnabledFormat);
            Assert.True(info.HasMacroProject, string.Join(" | ", info.Findings.Select(finding => finding.Message)));
            Assert.False(info.HasSignatures);
            Assert.NotNull(info.MacroProjectSha256);
            Assert.Equal(64, info.MacroProjectSha256!.Length);
        }

        [Fact]
        public void MacroSignatureInspectorParsesTheBoundedSnapshotThatPassedSecurityInspection() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureInspectionSnapshot.docm");
            string replacementPath = Path.Combine(_directoryWithFiles, "MacroSignatureInspectionReplacement.docm");
            using (WordDocument replacement = WordDocument.Create(replacementPath)) {
                replacement.AddParagraph("Replacement without a macro project");
                replacement.Save();
            }

            WordMacroProjectSignatureInfo info = WordMacroProjectSignatureInspector.Inspect(
                filePath,
                options: null,
                operationBudget: null,
                profileFilter: null,
                validateCmsOverride: null,
                afterSnapshotCaptured: () => File.Copy(replacementPath, filePath, overwrite: true));

            Assert.True(info.HasMacroProject, string.Join(" | ", info.Findings.Select(finding => finding.Message)));
            Assert.NotNull(info.MacroProjectPartUri);
            Assert.NotNull(info.MacroProjectSha256);
            Assert.False(WordDocument.InspectMacroProjectSignatures(filePath).HasMacroProject);
        }

        [Fact]
        public void MacroSignatureInspectorParsesAllProfilesAndCmsSigner() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureProfiles.docm");
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate("CN=OfficeIMO VBA Signing Test");
            AddMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.Legacy, certificate);
            AddMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.Agile, certificate);
            AddMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.V3, certificate);
            var options = CreateMacroSignatureInspectionOptions();

            WordMacroProjectSignatureInfo info = WordDocument.InspectMacroProjectSignatures(filePath, options);

            Assert.Equal(3, info.Signatures.Count);
            Assert.True(info.HasV3Signature);
            Assert.All(info.Signatures, signature => {
                Assert.True(signature.CmsParsed);
                Assert.Equal(WordSignatureValidationState.Passed, signature.CryptographicStatus);
                Assert.Equal(WordSignatureValidationState.Passed, signature.CertificateChainStatus);
                Assert.Equal(certificate.Thumbprint, signature.SignerThumbprint);
                Assert.Equal("2.16.840.1.101.3.4.2.1", signature.SignedContentDigestAlgorithmOid);
                Assert.Equal(32, signature.SignedContentDigest!.Length);
            });
        }

        [Fact]
        public void MacroSignatureInspectorRejectsOutOfRangeCmsBuffer() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureMalformed.docm");
            AddRawMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.V3,
                CreateMacroSignatureContainer(new byte[] { 0x30, 0x00 }, declaredLength: uint.MaxValue));

            WordMacroProjectSignatureInfo info = WordDocument.InspectMacroProjectSignatures(filePath);

            WordMacroProjectSignaturePartInfo signature = Assert.Single(info.Signatures);
            Assert.False(signature.CmsParsed);
            Assert.Equal(WordSignatureValidationState.Failed, signature.CryptographicStatus);
            Assert.Contains(signature.Findings, finding => finding.Code == "MacroSignatureContainerMalformed");
        }

        [Fact]
        public void MacroSignatureInspectorEnforcesAggregateSignatureBudget() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureBudget.docm");
            AddRawMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.Legacy, new byte[80]);
            AddRawMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.Agile, new byte[80]);
            var options = new WordMacroProjectSignatureInspectionOptions {
                ValidateCms = false,
                MaxSignatureBytes = 100,
                MaxTotalSignatureBytes = 100
            };
            WordMacroProjectSignatureInfo info = WordDocument.InspectMacroProjectSignatures(filePath, options);

            Assert.True(info.HasMacroProject);
            Assert.NotNull(info.MacroProjectPartUri);
            Assert.NotNull(info.MacroProjectSha256);
            Assert.Empty(info.Signatures);
            Assert.Contains(info.Findings, finding =>
                finding.Code == "MacroSignatureInspectionFailed" &&
                finding.Message.Contains("aggregate VBA signature bytes", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void MacroSignatureInspectorRetainsProjectIdentityWhenHashBudgetIsExceeded() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureProjectByteBudget.docm");
            var options = new WordMacroProjectSignatureInspectionOptions {
                MaxMacroProjectBytes = 1
            };

            WordMacroProjectSignatureInfo info = WordDocument.InspectMacroProjectSignatures(filePath, options);

            Assert.True(info.HasMacroProject);
            Assert.NotNull(info.MacroProjectPartUri);
            Assert.Null(info.MacroProjectLength);
            Assert.Null(info.MacroProjectSha256);
            Assert.Empty(info.Signatures);
            Assert.Contains(info.Findings, finding =>
                finding.Code == "MacroSignatureInspectionFailed" &&
                finding.Message.Contains("byte limit", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void MacroSignatureInspectorAppliesSharedPackageSecurityBeforeOpenXmlParsing() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignaturePackageBudget.docm");
            var options = new WordMacroProjectSignatureInspectionOptions();
            options.PackageSecurity.MaxPartCount = 1;

            WordMacroProjectSignatureInfo info = WordDocument.InspectMacroProjectSignatures(filePath, options);

            Assert.False(info.HasMacroProject);
            Assert.Contains(info.Findings, finding =>
                finding.Code == "PackageSecurityPartCount" &&
                finding.Message.Contains("before Open XML parsing", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void MacroSignatureValidationIsStructuredWhenPlatformIsUnsupported() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureUnsupported.docm");
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            AddMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.V3, certificate);
            var dependencies = new WordMacroProjectSigningDependencies(
                new RecordingMacroToolRunner(_ => throw new InvalidOperationException("Runner must not execute.")),
                new TestMacroSigningPlatform(isWindows: false));

            WordMacroProjectSignatureValidationResult result = WordMacroProjectSignatureService.Validate(
                filePath, new WordMacroProjectSignatureValidationOptions(), dependencies);

            Assert.False(result.IsSupported);
            Assert.False(result.IsValidUnderPolicy);
            Assert.Equal(WordSignatureValidationState.Unsupported, result.ContentBindingStatus);
            Assert.Contains(result.Findings, finding => finding.Code == "MacroSigningPlatformUnsupported");
        }

        [Fact]
        public void MacroSignatureValidationUsesOfficeSipAndCallerTrustPolicy() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureValidate.docm");
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            AddMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.V3, certificate);
            var runner = new RecordingMacroToolRunner(_ => Success());
            var dependencies = new WordMacroProjectSigningDependencies(
                runner, new TestMacroSigningPlatform(isWindows: true));
            var options = new WordMacroProjectSignatureValidationOptions();
            TrustMacroTestCertificate(options.Inspection.CmsVerification.CertificateValidation);

            WordMacroProjectSignatureValidationResult result = WordMacroProjectSignatureService.Validate(
                filePath, options, dependencies);

            Assert.True(result.IsSupported);
            Assert.True(result.IsValidUnderPolicy, string.Join(" | ", result.Findings.Select(
                finding => finding.Code + ": " + finding.Message)));
            Assert.Equal(WordSignatureValidationState.Passed, result.ContentBindingStatus);
            Assert.Empty(runner.Invocations);
        }

        [Fact]
        public void MacroSignatureValidationUsesOneImmutablePackageSnapshot() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureValidationSnapshot.docm");
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            AddMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.V3, certificate);
            byte[] replacement = File.ReadAllBytes(CreateMacroEnabledTestDocument("MacroSignatureReplacement.docm"));
            var platform = new SnapshotRecordingMacroSigningPlatform(() => File.WriteAllBytes(filePath, replacement));
            var dependencies = new WordMacroProjectSigningDependencies(
                new RecordingMacroToolRunner(_ => Success()), platform);
            var options = new WordMacroProjectSignatureValidationOptions();
            TrustMacroTestCertificate(options.Inspection.CmsVerification.CertificateValidation);

            WordMacroProjectSignatureValidationResult result = WordMacroProjectSignatureService.Validate(
                filePath, options, dependencies);

            Assert.True(result.IsValidUnderPolicy, string.Join(" | ", result.Findings.Select(
                finding => finding.Code + ": " + finding.Message)));
            Assert.Equal(Path.GetFullPath(filePath), result.SignatureInfo.FilePath);
            Assert.Equal(2, platform.Paths.Count);
            Assert.Equal(platform.Paths[0], platform.Paths[1]);
            Assert.NotEqual(Path.GetFullPath(filePath), platform.Paths[0]);
        }

        [Fact]
        public void MacroSignatureValidationAcceptsValidHighestProfileWithMalformedLowerProfile() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureProfilePrecedence.docm");
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            AddRawMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.Legacy, new byte[48]);
            AddMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.V3, certificate);
            var options = new WordMacroProjectSignatureValidationOptions();
            TrustMacroTestCertificate(options.Inspection.CmsVerification.CertificateValidation);
            var dependencies = new WordMacroProjectSigningDependencies(
                new RecordingMacroToolRunner(_ => Success()),
                new TestMacroSigningPlatform(isWindows: true));

            WordMacroProjectSignatureValidationResult result = WordMacroProjectSignatureService.Validate(
                filePath, options, dependencies);

            Assert.True(result.IsValidUnderPolicy, string.Join(" | ", result.Findings.Select(
                finding => finding.Code + ": " + finding.Message)));
            Assert.Contains(result.Findings, finding =>
                finding.Profile == WordMacroProjectSignatureProfile.Legacy &&
                finding.Code == "MacroSignatureContainerMalformed");
        }

        [Fact]
        public void MacroSignatureValidationEnforcesV3AndTimestampPolicyBeforeAcceptance() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignaturePolicy.docm");
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            AddMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.Agile, certificate);
            var runner = new RecordingMacroToolRunner(_ => Success());
            var dependencies = new WordMacroProjectSigningDependencies(
                runner, new TestMacroSigningPlatform(isWindows: true));
            var options = new WordMacroProjectSignatureValidationOptions();
            TrustMacroTestCertificate(options.Inspection.CmsVerification.CertificateValidation);

            WordMacroProjectSignatureValidationResult missingV3 = WordMacroProjectSignatureService.Validate(
                filePath, options, dependencies);

            Assert.False(missingV3.IsValidUnderPolicy);
            Assert.Equal(WordSignatureValidationState.NotChecked, missingV3.ContentBindingStatus);
            Assert.Empty(runner.Invocations);
            Assert.Contains(missingV3.Findings, finding => finding.Code == "MacroV3SignatureRequired");

            options.RequireV3Signature = false;
            options.RequireTimestamp = true;
            WordMacroProjectSignatureValidationResult missingTimestamp = WordMacroProjectSignatureService.Validate(
                filePath, options, dependencies);

            Assert.Equal(WordSignatureValidationState.Passed, missingTimestamp.ContentBindingStatus);
            Assert.False(missingTimestamp.IsValidUnderPolicy);
            Assert.Empty(runner.Invocations);

            options.Inspection.ValidateCms = false;
            Assert.Throws<ArgumentException>(() => WordMacroProjectSignatureService.Validate(
                filePath, options, dependencies));
        }

        [Fact]
        public void MacroSigningCreatesThreeProfilesAndCommitsAtomically() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureAtomic.docm");
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            string originalMacroHash = WordDocument.InspectMacroProjectSignatures(filePath).MacroProjectSha256!;
            string toolsDirectory = CreateFakeOfficeSipsDirectory();
            var runner = new SimulatedOfficeSipsRunner(certificate);
            var dependencies = new WordMacroProjectSigningDependencies(
                runner, new TestMacroSigningPlatform(isWindows: true));
            var options = new WordMacroProjectSigningOptions {
                OfficeSipsDirectory = toolsDirectory,
                TimestampAuthorityUrl = new Uri("https://timestamp.example.test/rfc3161"),
                StoreName = StoreName.CertificateAuthority
            };
            TrustMacroTestCertificate(options.Inspection.CmsVerification.CertificateValidation);

            WordMacroProjectSigningResult result = WordMacroProjectSignatureService.TrySign(
                filePath, certificate.Thumbprint!, options, dependencies);

            Assert.True(result.Succeeded);
            Assert.True(result.MacroProjectPreserved);
            Assert.NotNull(result.ValidationResult);
            Assert.True(result.ValidationResult!.IsValidUnderPolicy);
            Assert.Equal(4, runner.Invocations.Count);
            Assert.EndsWith("offclearsig.exe", runner.Invocations[0].ExecutablePath, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(new[] { "sign", "sign", "sign" },
                runner.Invocations.Skip(1).Select(invocation => invocation.Arguments[0]).ToArray());
            Assert.All(runner.Invocations.Where(invocation => invocation.Arguments[0] == "sign"), invocation => {
                Assert.Contains("/sha1", invocation.Arguments);
                Assert.Contains(certificate.Thumbprint!, invocation.Arguments);
                Assert.Contains("/fd", invocation.Arguments);
                Assert.Contains("SHA256", invocation.Arguments);
                int storeIndex = invocation.Arguments.ToList().IndexOf("/s");
                Assert.True(storeIndex >= 0);
                Assert.Equal("CA", invocation.Arguments[storeIndex + 1]);
                Assert.Contains("/tr", invocation.Arguments);
                Assert.DoesNotContain("/f", invocation.Arguments);
                Assert.DoesNotContain("/as", invocation.Arguments);
            });
            WordMacroProjectSignatureInfo committed = WordDocument.InspectMacroProjectSignatures(filePath, options.Inspection);
            Assert.Equal(3, committed.Signatures.Count);
            Assert.Equal(originalMacroHash, committed.MacroProjectSha256);
        }

        [Fact]
        public void MacroSigningReadbackFailurePreservesDetailedInspectionEvidence() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureReadbackDiagnostics.docm");
            string toolsDirectory = CreateFakeOfficeSipsDirectory();
            var runner = new RecordingMacroToolRunner(invocation => {
                if (invocation.Arguments.Count > 0 && invocation.Arguments[0] == "sign") {
                    AddRawMacroSignatureProfile(
                        invocation.Arguments[invocation.Arguments.Count - 1],
                        WordMacroProjectSignatureProfile.Legacy,
                        CreateMacroSignatureContainer(new byte[] { 0x30, 0x00 }));
                }
                return Success();
            });
            var dependencies = new WordMacroProjectSigningDependencies(
                runner, new TestMacroSigningPlatform(isWindows: true));
            var options = new WordMacroProjectSigningOptions { OfficeSipsDirectory = toolsDirectory };

            WordMacroProjectSigningResult result = WordMacroProjectSignatureService.TrySign(
                filePath, "00112233445566778899AABBCCDDEEFF00112233", options, dependencies);

            Assert.False(result.Succeeded);
            Assert.Contains(result.Findings, finding => finding.Code == "CmsMalformed");
            Assert.Contains(result.Findings, finding => finding.Code == "MacroAgileReadbackFailed");
        }

        [Fact]
        public void MacroSigningBoundsTheSourceSnapshotCopy() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureBoundedSource.docm");
            long originalLength = new FileInfo(filePath).Length;
            string toolsDirectory = CreateFakeOfficeSipsDirectory();
            var runner = new RecordingMacroToolRunner(_ => Success());
            var platform = new SourceMutationMacroSigningPlatform(() =>
                File.WriteAllBytes(filePath, new byte[checked((int)originalLength + 4096)]));
            var dependencies = new WordMacroProjectSigningDependencies(runner, platform);
            var options = new WordMacroProjectSigningOptions { OfficeSipsDirectory = toolsDirectory };
            options.Inspection.PackageSecurity.MaxPackageBytes = originalLength + 1024;

            WordMacroProjectSigningResult result = WordMacroProjectSignatureService.TrySign(
                filePath, "00112233445566778899AABBCCDDEEFF00112233", options, dependencies);

            Assert.False(result.Succeeded);
            Assert.Empty(runner.Invocations);
            Assert.Contains(result.Findings, finding =>
                finding.Code == "MacroProjectSigningFailed" &&
                finding.Message.Contains("configured limit", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void MacroSigningFailurePreservesOriginalAndDeletesStage() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureRollback.docm");
            byte[] originalBytes = File.ReadAllBytes(filePath);
            string toolsDirectory = CreateFakeOfficeSipsDirectory();
            var runner = new RecordingMacroToolRunner(invocation =>
                invocation.ExecutablePath.EndsWith("offclearsig.exe", StringComparison.OrdinalIgnoreCase)
                    ? Success()
                    : new WordMacroProjectToolResult(9, false, "simulated signing failure"));
            var dependencies = new WordMacroProjectSigningDependencies(
                runner, new TestMacroSigningPlatform(isWindows: true));
            var options = new WordMacroProjectSigningOptions { OfficeSipsDirectory = toolsDirectory };

            WordMacroProjectSigningResult result = WordMacroProjectSignatureService.TrySign(
                filePath, "00112233445566778899AABBCCDDEEFF00112233", options, dependencies);

            Assert.False(result.Succeeded);
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
            string prefix = "." + Path.GetFileNameWithoutExtension(filePath) + ".";
            Assert.DoesNotContain(Directory.EnumerateFiles(Path.GetDirectoryName(filePath)!, "*.docm"),
                path => Path.GetFileName(path).StartsWith(prefix, StringComparison.Ordinal));
        }

        [Fact]
        public void MacroSigningBlocksExistingOpcSignaturesUnlessInvalidationIsExplicit() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureOpcPolicy.docm");
            AddDigitalSignatureMetadata(filePath, CreateSignatureXml());
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            string toolsDirectory = CreateFakeOfficeSipsDirectory();
            var blockedRunner = new SimulatedOfficeSipsRunner(certificate);
            var dependencies = new WordMacroProjectSigningDependencies(
                blockedRunner, new TestMacroSigningPlatform(isWindows: true));
            var options = new WordMacroProjectSigningOptions { OfficeSipsDirectory = toolsDirectory };
            TrustMacroTestCertificate(options.Inspection.CmsVerification.CertificateValidation);

            WordMacroProjectSigningResult blocked = WordMacroProjectSignatureService.TrySign(
                filePath, certificate.Thumbprint!, options, dependencies);

            Assert.False(blocked.Succeeded);
            Assert.Empty(blockedRunner.Invocations);
            Assert.Contains(blocked.Findings,
                finding => finding.Code == "ExistingPackageSignatureInvalidationBlocked");

            options.ExistingPackageSignaturePolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation;
            WordMacroProjectSigningResult allowed = WordMacroProjectSignatureService.TrySign(
                filePath, certificate.Thumbprint!, options, dependencies);

            Assert.True(allowed.Succeeded);
            Assert.Contains(allowed.Findings,
                finding => finding.Code == "ExistingPackageSignaturesWillBeInvalidated");
        }

        [Fact]
        public void MacroSigningPreflightsTheExactStagedSnapshotForOpcSignatures() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureOpcSnapshot.docm");
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            string toolsDirectory = CreateFakeOfficeSipsDirectory();
            var runner = new SimulatedOfficeSipsRunner(certificate);
            var platform = new SourceMutationMacroSigningPlatform(() =>
                AddDigitalSignatureMetadata(filePath, CreateSignatureXml()));
            var dependencies = new WordMacroProjectSigningDependencies(runner, platform);
            var options = new WordMacroProjectSigningOptions { OfficeSipsDirectory = toolsDirectory };
            TrustMacroTestCertificate(options.Inspection.CmsVerification.CertificateValidation);

            WordMacroProjectSigningResult result = WordMacroProjectSignatureService.TrySign(
                filePath, certificate.Thumbprint!, options, dependencies);

            Assert.False(result.Succeeded);
            Assert.Empty(runner.Invocations);
            Assert.Contains(result.Findings,
                finding => finding.Code == "ExistingPackageSignatureInvalidationBlocked");
        }

        [Fact]
        public void MacroSigningDoesNotOverwriteAConcurrentSourceChange() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureConcurrent.docm");
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            string toolsDirectory = CreateFakeOfficeSipsDirectory();
            var runner = new ConcurrentSourceMutationRunner(certificate, filePath);
            var dependencies = new WordMacroProjectSigningDependencies(
                runner, new TestMacroSigningPlatform(isWindows: true));
            var options = new WordMacroProjectSigningOptions { OfficeSipsDirectory = toolsDirectory };
            TrustMacroTestCertificate(options.Inspection.CmsVerification.CertificateValidation);

            WordMacroProjectSigningResult result = WordMacroProjectSignatureService.TrySign(
                filePath, certificate.Thumbprint!, options, dependencies);

            Assert.False(result.Succeeded);
            Assert.Contains(result.Findings, finding => finding.Code == "SourcePackageChangedDuringSigning");
            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            Assert.True(package.PartExists(new Uri("/word/concurrent-change.bin", UriKind.Relative)));
        }

        [Fact]
        public void MacroSigningRejectsArbitraryCertificateSelectionTextBeforeToolExecution() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureThumbprint.docm");
            var runner = new RecordingMacroToolRunner(_ => Success());
            var dependencies = new WordMacroProjectSigningDependencies(
                runner, new TestMacroSigningPlatform(isWindows: true));

            WordMacroProjectSigningResult result = WordMacroProjectSignatureService.TrySign(
                filePath, "0011 & calc.exe", new WordMacroProjectSigningOptions(), dependencies);

            Assert.False(result.Succeeded);
            Assert.Empty(runner.Invocations);
            Assert.Contains(result.Findings, finding => finding.Code == "CertificateThumbprintInvalid");
        }

        [Fact]
        public void MacroSigningRequiresTimestampAuthorityWhenTimestampIsMandatory() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureTimestampPolicy.docm");
            var dependencies = new WordMacroProjectSigningDependencies(
                new RecordingMacroToolRunner(_ => Success()),
                new TestMacroSigningPlatform(isWindows: true));
            var options = new WordMacroProjectSigningOptions { RequireTimestamp = true };

            Assert.Throws<ArgumentException>(() => WordMacroProjectSignatureService.TrySign(
                filePath, "00112233445566778899AABBCCDDEEFF00112233", options, dependencies));
        }

        [Fact]
        public void PackageAndMacroSigningRemainSeparateCapabilities() {
            Assert.Equal(WordSigningCapabilityKind.OpcPackage, WordSigningCapabilities.Package.Kind);
            Assert.Equal(WordSigningCapabilityKind.VbaMacroProject, WordSigningCapabilities.MacroProject.Kind);
            Assert.True(WordSigningCapabilities.Package.IsSupported);
            Assert.Contains("separate", WordSigningCapabilities.MacroProject.Message, StringComparison.OrdinalIgnoreCase);
        }

        [VbaOfficeSipInteropFact]
        public void MacroSigningInteroperatesWithMicrosoftOfficeSipAndDetectsTampering() {
            TraceOfficeSipInterop("document-create:start");
            string sourcePath = Environment.GetEnvironmentVariable("OFFICEIMO_VBA_INTEROP_DOCUMENT_PATH")!;
            string filePath = Path.Combine(_directoryWithFiles, "MacroSignatureOfficeSips.docm");
            File.Copy(sourcePath, filePath, overwrite: true);
            TraceOfficeSipInterop("document-create:complete");

            TraceOfficeSipInterop("certificate-create:start");
            using X509Certificate2 certificate = CreateOfficeSipSigningCertificate();
            TraceOfficeSipInterop("certificate-create:complete");
            using var personalStore = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            TraceOfficeSipInterop("certificate-store-open:start");
            personalStore.Open(OpenFlags.ReadWrite);
            TraceOfficeSipInterop("certificate-store-open:complete");
            TraceOfficeSipInterop("certificate-personal-store-add:start");
            personalStore.Add(certificate);
            TraceOfficeSipInterop("certificate-personal-store-add:complete");
            using var trustedPeopleStore = new X509Store(StoreName.TrustedPeople, StoreLocation.CurrentUser);
            TraceOfficeSipInterop("certificate-trusted-people-store-open:start");
            trustedPeopleStore.Open(OpenFlags.ReadWrite);
            TraceOfficeSipInterop("certificate-trusted-people-store-open:complete");
            TraceOfficeSipInterop("certificate-trusted-people-store-add:start");
            trustedPeopleStore.Add(certificate);
            TraceOfficeSipInterop("certificate-trusted-people-store-add:complete");
            try {
                var options = new WordMacroProjectSigningOptions {
                    SignToolPath = Environment.GetEnvironmentVariable("OFFICEIMO_SIGNTOOL_PATH"),
                    OfficeSipsDirectory = Environment.GetEnvironmentVariable("OFFICEIMO_VBA_SIP_DIRECTORY"),
                    ToolTimeout = TimeSpan.FromSeconds(30)
                };
                TrustMacroTestCertificate(options.Inspection.CmsVerification.CertificateValidation);
                var dependencies = new WordMacroProjectSigningDependencies(
                    new TracingOfficeSipToolRunner(new WordMacroProjectProcessRunner()),
                    new TracingOfficeSipPlatform(new WordMacroProjectWindowsPlatform()));

                TraceOfficeSipInterop("macro-sign:start");
                WordMacroProjectSigningResult signing = WordMacroProjectSignatureService.TrySign(
                    filePath, certificate.Thumbprint!, options, dependencies);
                TraceOfficeSipInterop("macro-sign:complete:" + signing.Succeeded);

                Assert.True(signing.Succeeded, string.Join(" | ", signing.Findings.Select(
                    finding => finding.Code + ": " + finding.Message)));
                Assert.True(signing.MacroProjectPreserved);
                Assert.True(signing.ValidationResult!.IsValidUnderPolicy);
                Assert.Equal(3, signing.ValidationResult.SignatureInfo.Signatures.Count);

                TraceOfficeSipInterop("macro-validate:start");
                WordMacroProjectSignatureValidationResult valid =
                    WordMacroProjectSignatureService.Validate(filePath, options, dependencies);
                TraceOfficeSipInterop("macro-validate:complete:" + valid.IsValidUnderPolicy);
                Assert.True(valid.IsValidUnderPolicy);

                TraceOfficeSipInterop("macro-tamper:start");
                TamperVbaProjectPart(filePath);
                TraceOfficeSipInterop("macro-tamper:complete");
                TraceOfficeSipInterop("tampered-validate:start");
                WordMacroProjectSignatureValidationResult tampered =
                    WordMacroProjectSignatureService.Validate(filePath, options, dependencies);
                TraceOfficeSipInterop("tampered-validate:complete:" + tampered.IsValidUnderPolicy);
                Assert.False(tampered.IsValidUnderPolicy);
                Assert.Equal(WordSignatureValidationState.Failed, tampered.ContentBindingStatus);
            } finally {
                TraceOfficeSipInterop("certificate-store-remove:start");
                RemoveCertificatesByThumbprint(trustedPeopleStore, certificate.Thumbprint);
                RemoveCertificatesByThumbprint(personalStore, certificate.Thumbprint);
                TraceOfficeSipInterop("certificate-store-remove:complete");
            }
        }

        private string CreateMacroEnabledTestDocument(string fileName) {
            string vbaPath = Path.Combine(_directoryWithFiles, Path.GetFileNameWithoutExtension(fileName) + ".vba.bin");
            string filePath = Path.Combine(_directoryWithFiles, fileName);
            CreateDummyVba(vbaPath, "Module1");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddMacro(vbaPath);
                document.Save();
            }
            return filePath;
        }

        private string CreateFakeOfficeSipsDirectory() {
            string directory = Path.Combine(_directoryWithFiles, "OfficeSips-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, "offclearsig.exe"), new byte[] { 0x4D, 0x5A });
            return directory;
        }

        private static WordMacroProjectSignatureInspectionOptions CreateMacroSignatureInspectionOptions() {
            var options = new WordMacroProjectSignatureInspectionOptions();
            TrustMacroTestCertificate(options.CmsVerification.CertificateValidation);
            return options;
        }

        private static void TrustMacroTestCertificate(
            OfficeIMO.Security.CertificateValidationOptions options) {
            options.ChainEvaluator = (_, _) => true;
            options.DisableCertificateDownloads = false;
        }

        private static void AddMacroSignatureProfile(
            string filePath,
            WordMacroProjectSignatureProfile profile,
            X509Certificate2 certificate) {
            byte[] cms = CreateAuthenticodeCms(profile, certificate);
            AddRawMacroSignatureProfile(filePath, profile, CreateMacroSignatureContainer(cms));
        }

        private static byte[] CreateAuthenticodeCms(
            WordMacroProjectSignatureProfile profile,
            X509Certificate2 certificate) {
            byte[] digestInput = System.Text.Encoding.UTF8.GetBytes("OfficeIMO VBA binding " + profile);
            byte[] subjectDigest;
            using (System.Security.Cryptography.SHA256 sha256 = System.Security.Cryptography.SHA256.Create()) {
                subjectDigest = sha256.ComputeHash(digestInput);
            }
            var digestAlgorithm = new Org.BouncyCastle.Asn1.X509.AlgorithmIdentifier(
                Org.BouncyCastle.Asn1.Nist.NistObjectIdentifiers.IdSha256,
                Org.BouncyCastle.Asn1.DerNull.Instance);
            var indirectData = new Org.BouncyCastle.Asn1.DerSequence(
                new Org.BouncyCastle.Asn1.DerSequence(
                    new Org.BouncyCastle.Asn1.DerObjectIdentifier("1.3.6.1.4.1.311.2.1.15")),
                new Org.BouncyCastle.Asn1.X509.DigestInfo(digestAlgorithm, subjectDigest));
            byte[] encoded = OfficeIMO.Security.CmsSignedDataSigner.SignEncapsulated(
                indirectData.GetEncoded(),
                certificate,
                new OfficeIMO.Security.CmsSigningOptions {
                    ContentTypeOid = "1.3.6.1.4.1.311.2.1.4",
                    IncludeCertificateChain = false
                });
            var decoded = new Org.BouncyCastle.Cms.CmsSignedData(encoded);
            Assert.Equal("1.3.6.1.4.1.311.2.1.4", decoded.SignedContentType.Id);
            return encoded;
        }

        private static byte[] CreateMacroSignatureContainer(byte[] cms, uint? declaredLength = null) {
            const int digSigBlobHeaderLength = 8;
            const int digSigInfoSerializedHeaderLength = 36;
            byte[] bytes = new byte[digSigInfoSerializedHeaderLength + cms.Length];
            WriteUInt32(bytes, 0, declaredLength ?? checked((uint)cms.Length));
            WriteUInt32(bytes, 4, digSigBlobHeaderLength + digSigInfoSerializedHeaderLength);
            Buffer.BlockCopy(cms, 0, bytes, digSigInfoSerializedHeaderLength, cms.Length);
            return bytes;
        }

        private static void AddRawMacroSignatureProfile(
            string filePath,
            WordMacroProjectSignatureProfile profile,
            byte[] bytes) {
            Uri vbaUri;
            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, false)) {
                vbaUri = document.MainDocumentPart!.VbaProjectPart!.Uri;
            }
            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite);
            PackagePart vbaPart = package.GetPart(vbaUri);
            string partName;
            string contentType;
            string relationshipType;
            switch (profile) {
                case WordMacroProjectSignatureProfile.Legacy:
                    partName = "vbaProjectSignature.bin";
                    contentType = "application/vnd.ms-office.vbaProjectSignature";
                    relationshipType = WordMacroProjectSignatureInspector.LegacyRelationship;
                    break;
                case WordMacroProjectSignatureProfile.Agile:
                    partName = "vbaProjectSignatureAgile.bin";
                    contentType = "application/vnd.ms-office.vbaProjectSignatureAgile";
                    relationshipType = WordMacroProjectSignatureInspector.AgileRelationship;
                    break;
                case WordMacroProjectSignatureProfile.V3:
                    partName = "vbaProjectSignatureV3.bin";
                    contentType = "application/vnd.ms-office.vbaProjectSignatureV3";
                    relationshipType = WordMacroProjectSignatureInspector.V3Relationship;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(profile));
            }
            Uri partUri = new Uri("/word/" + partName, UriKind.Relative);
            if (package.PartExists(partUri)) package.DeletePart(partUri);
            PackagePart signaturePart = package.CreatePart(partUri, contentType);
            using (Stream stream = signaturePart.GetStream(FileMode.Create, FileAccess.Write)) {
                stream.Write(bytes, 0, bytes.Length);
            }
            foreach (PackageRelationship relationship in vbaPart.GetRelationshipsByType(relationshipType).ToArray()) {
                vbaPart.DeleteRelationship(relationship.Id);
            }
            vbaPart.CreateRelationship(new Uri(partName, UriKind.Relative), TargetMode.Internal, relationshipType);
        }

        private static void WriteUInt32(byte[] bytes, int offset, uint value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }

        private static X509Certificate2 CreateOfficeSipSigningCertificate() {
            using System.Security.Cryptography.RSA rsa = System.Security.Cryptography.RSA.Create(2048);
            var request = new System.Security.Cryptography.X509Certificates.CertificateRequest(
                "CN=OfficeIMO OfficeSips Interop",
                rsa,
                System.Security.Cryptography.HashAlgorithmName.SHA256,
                System.Security.Cryptography.RSASignaturePadding.Pkcs1);
            request.CertificateExtensions.Add(new X509KeyUsageExtension(
                X509KeyUsageFlags.DigitalSignature, critical: true));
            request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(
                new System.Security.Cryptography.OidCollection {
                    new System.Security.Cryptography.Oid("1.3.6.1.5.5.7.3.3")
                }, critical: false));
            using X509Certificate2 created = request.CreateSelfSigned(
                DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
            return new X509Certificate2(created.Export(X509ContentType.Pfx), (string?)null,
                X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.UserKeySet);
        }

        private static void RemoveCertificatesByThumbprint(X509Store store, string? thumbprint) {
            if (string.IsNullOrWhiteSpace(thumbprint)) return;
            X509Certificate2Collection matches = store.Certificates.Find(
                X509FindType.FindByThumbprint, thumbprint, validOnly: false);
            foreach (X509Certificate2 match in matches) store.Remove(match);
        }

        private static void TamperVbaProjectPart(string filePath) {
            Uri vbaUri;
            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, false)) {
                vbaUri = document.MainDocumentPart!.VbaProjectPart!.Uri;
            }
            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite);
            PackagePart part = package.GetPart(vbaUri);
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.ReadWrite);
            byte[] bytes;
            using (var copy = new MemoryStream()) {
                stream.CopyTo(copy);
                bytes = copy.ToArray();
            }
            Assert.NotEmpty(bytes);
            bytes[bytes.Length - 1] ^= 0x01;
            stream.Position = 0;
            stream.SetLength(0);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static WordMacroProjectToolResult Success() =>
            new WordMacroProjectToolResult(0, false, "success");

        private sealed class TestMacroSigningPlatform : IWordMacroProjectPlatform {
            internal TestMacroSigningPlatform(bool isWindows) => IsWindows = isWindows;
            public bool IsWindows { get; }
            public bool TryGetSubjectInterfacePackage(string filePath, out Guid subjectGuid, out string detail) {
                subjectGuid = new Guid("6E64D5BD-CEB0-4B66-B4A0-15AC71775C48");
                detail = "simulated Microsoft Office SIP";
                return IsWindows;
            }
            public WordMacroProjectContentBindingResult ValidateContentBinding(
                string filePath,
                string digestAlgorithmOid,
                byte[] expectedDigest) =>
                IsWindows
                    ? new WordMacroProjectContentBindingResult(true, true, "simulated Office SIP digest match")
                    : new WordMacroProjectContentBindingResult(false, false, "simulated unsupported platform");
        }

        private sealed class SourceMutationMacroSigningPlatform : IWordMacroProjectPlatform {
            private readonly Action _mutateOnce;
            private bool _mutated;

            internal SourceMutationMacroSigningPlatform(Action mutateOnce) => _mutateOnce = mutateOnce;

            public bool IsWindows => true;

            public bool TryGetSubjectInterfacePackage(string filePath, out Guid subjectGuid, out string detail) {
                if (!_mutated) {
                    _mutated = true;
                    _mutateOnce();
                }
                subjectGuid = new Guid("6E64D5BD-CEB0-4B66-B4A0-15AC71775C48");
                detail = "simulated Microsoft Office SIP";
                return true;
            }

            public WordMacroProjectContentBindingResult ValidateContentBinding(
                string filePath,
                string digestAlgorithmOid,
                byte[] expectedDigest) =>
                new WordMacroProjectContentBindingResult(true, true, "simulated Office SIP digest match");
        }

        private sealed class SnapshotRecordingMacroSigningPlatform : IWordMacroProjectPlatform {
            private readonly Action _mutateOriginal;

            internal SnapshotRecordingMacroSigningPlatform(Action mutateOriginal) =>
                _mutateOriginal = mutateOriginal;

            internal List<string> Paths { get; } = new List<string>();
            public bool IsWindows => true;

            public bool TryGetSubjectInterfacePackage(string filePath, out Guid subjectGuid, out string detail) {
                Paths.Add(filePath);
                _mutateOriginal();
                subjectGuid = new Guid("6E64D5BD-CEB0-4B66-B4A0-15AC71775C48");
                detail = "simulated Microsoft Office SIP";
                return true;
            }

            public WordMacroProjectContentBindingResult ValidateContentBinding(
                string filePath,
                string digestAlgorithmOid,
                byte[] expectedDigest) {
                Paths.Add(filePath);
                return new WordMacroProjectContentBindingResult(true, true,
                    "simulated Office SIP digest match on immutable snapshot");
            }
        }

        private sealed class RecordingMacroToolRunner : IWordMacroProjectToolRunner {
            private readonly Func<WordMacroProjectToolInvocation, WordMacroProjectToolResult> _run;
            internal RecordingMacroToolRunner(Func<WordMacroProjectToolInvocation, WordMacroProjectToolResult> run) => _run = run;
            internal List<WordMacroProjectToolInvocation> Invocations { get; } = new List<WordMacroProjectToolInvocation>();
            public WordMacroProjectToolResult Run(WordMacroProjectToolInvocation invocation, TimeSpan timeout,
                int maxOutputCharacters) {
                Invocations.Add(invocation);
                return _run(invocation);
            }
        }

        private sealed class SimulatedOfficeSipsRunner : IWordMacroProjectToolRunner {
            private readonly X509Certificate2 _certificate;
            private readonly bool _includeTimestampToken;
            private int _profile;
            internal SimulatedOfficeSipsRunner(X509Certificate2 certificate, bool includeTimestampToken = false) {
                _certificate = certificate;
                _includeTimestampToken = includeTimestampToken;
            }
            internal List<WordMacroProjectToolInvocation> Invocations { get; } = new List<WordMacroProjectToolInvocation>();

            public WordMacroProjectToolResult Run(WordMacroProjectToolInvocation invocation, TimeSpan timeout,
                int maxOutputCharacters) {
                Invocations.Add(invocation);
                if (invocation.Arguments.Count > 0 && invocation.Arguments[0] == "sign") {
                    _profile++;
                    if (_includeTimestampToken) {
                        AddTimestampedMacroSignatureProfile(invocation.Arguments[invocation.Arguments.Count - 1],
                            (WordMacroProjectSignatureProfile)_profile, _certificate);
                    } else {
                        AddMacroSignatureProfile(invocation.Arguments[invocation.Arguments.Count - 1],
                            (WordMacroProjectSignatureProfile)_profile, _certificate);
                    }
                }
                return Success();
            }
        }

        private sealed class ConcurrentSourceMutationRunner : IWordMacroProjectToolRunner {
            private readonly SimulatedOfficeSipsRunner _inner;
            private readonly string _sourcePath;
            private bool _mutated;

            internal ConcurrentSourceMutationRunner(X509Certificate2 certificate, string sourcePath) {
                _inner = new SimulatedOfficeSipsRunner(certificate);
                _sourcePath = sourcePath;
            }

            public WordMacroProjectToolResult Run(WordMacroProjectToolInvocation invocation, TimeSpan timeout,
                int maxOutputCharacters) {
                if (!_mutated && invocation.Arguments.Count > 0 && invocation.Arguments[0] == "sign") {
                    _mutated = true;
                    using Package package = Package.Open(_sourcePath, FileMode.Open, FileAccess.ReadWrite);
                    PackagePart part = package.CreatePart(
                        new Uri("/word/concurrent-change.bin", UriKind.Relative), "application/octet-stream");
                    using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
                    stream.WriteByte(0x42);
                }
                return _inner.Run(invocation, timeout, maxOutputCharacters);
            }
        }
    }
}

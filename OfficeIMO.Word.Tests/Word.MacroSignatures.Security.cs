using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void MacroSignatureInspectorSharesTimestampCountBudgetAcrossProfiles() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureTimestampCountBudget.docm");
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            AddTimestampedMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.Legacy, certificate);
            AddTimestampedMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.Agile, certificate);
            AddTimestampedMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.V3, certificate);
            var options = CreateMacroSignatureInspectionOptions();
            options.CmsVerification.MaxTimestampTokens = 2;

            WordMacroProjectSignatureInfo info = WordDocument.InspectMacroProjectSignatures(filePath, options);

            Assert.Equal(3, info.Signatures.Count);
            Assert.Contains(info.Signatures.Single(signature =>
                    signature.Profile == WordMacroProjectSignatureProfile.V3).Findings,
                finding => finding.Code == "CmsTimestampCountLimitExceeded");

            options.CmsVerification.MaxTimestampTokens = 3;
            WordMacroProjectSignatureInfo permissive = WordDocument.InspectMacroProjectSignatures(filePath, options);
            Assert.DoesNotContain(permissive.Signatures.SelectMany(signature => signature.Findings),
                finding => finding.Code == "CmsTimestampCountLimitExceeded");
        }

        [Fact]
        public void MacroSignatureInspectorSharesTimestampByteBudgetAcrossProfiles() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureTimestampByteBudget.docm");
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            AddTimestampedMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.Legacy, certificate);
            AddTimestampedMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.Agile, certificate);
            var options = CreateMacroSignatureInspectionOptions();
            options.CmsVerification.MaxTimestampTokenBytes = 16;
            options.CmsVerification.MaxTotalTimestampBytes = 3;

            WordMacroProjectSignatureInfo info = WordDocument.InspectMacroProjectSignatures(filePath, options);

            Assert.Contains(info.Signatures.Single(signature =>
                    signature.Profile == WordMacroProjectSignatureProfile.Agile).Findings,
                finding => finding.Code == "CmsTimestampTotalSizeLimitExceeded");

            options.CmsVerification.MaxTotalTimestampBytes = 4;
            WordMacroProjectSignatureInfo permissive = WordDocument.InspectMacroProjectSignatures(filePath, options);
            Assert.DoesNotContain(permissive.Signatures.SelectMany(signature => signature.Findings),
                finding => finding.Code == "CmsTimestampTotalSizeLimitExceeded");
        }

        [Fact]
        public void MacroSignatureInspectorCarriesTimestampBudgetAcrossSigningReadbacks() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureTimestampReadbackBudget.docm");
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            AddTimestampedMacroSignatureProfile(filePath, WordMacroProjectSignatureProfile.V3, certificate);
            var options = CreateMacroSignatureInspectionOptions();
            options.CmsVerification.MaxTimestampTokens = 1;
            var budget = new WordMacroProjectSignatureInspector.InspectionBudget(options);

            WordMacroProjectSignatureInfo first = WordMacroProjectSignatureInspector.Inspect(
                filePath, options, budget, WordMacroProjectSignatureProfile.V3);
            WordMacroProjectSignatureInfo second = WordMacroProjectSignatureInspector.Inspect(
                filePath, options, budget, WordMacroProjectSignatureProfile.V3);

            Assert.DoesNotContain(first.Signatures.SelectMany(signature => signature.Findings),
                finding => finding.Code == "CmsTimestampCountLimitExceeded");
            Assert.Contains(second.Signatures.SelectMany(signature => signature.Findings),
                finding => finding.Code == "CmsTimestampCountLimitExceeded");
        }

        [Fact]
        public void MacroSigningRejectsReplacementOfTheValidatedStagingPath() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureReplacedStage.docm");
            byte[] originalBytes = File.ReadAllBytes(filePath);
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            string toolsDirectory = CreateFakeOfficeSipsDirectory();
            var runner = new SimulatedOfficeSipsRunner(certificate);
            var platform = new StagingReplacementMacroSigningPlatform(runner);
            var dependencies = new WordMacroProjectSigningDependencies(runner, platform);
            var options = new WordMacroProjectSigningOptions { OfficeSipsDirectory = toolsDirectory };
            options.Inspection.CmsVerification.CertificateValidation.ChainEvaluator = (_, _) => true;

            WordMacroProjectSigningResult result = WordMacroProjectSignatureService.TrySign(
                filePath, certificate.Thumbprint!, options, dependencies);

            Assert.True(platform.ReplacedStagingPath);
            Assert.False(result.Succeeded);
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
            Assert.Contains(result.Findings, finding => finding.Code == "SourcePackageChangedDuringSigning");
        }

        [Fact]
        public void MacroSigningDoesNotClaimPreservationWhenValidationSnapshotChanges() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureChangedValidationSnapshot.docm");
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            string toolsDirectory = CreateFakeOfficeSipsDirectory();
            var runner = new SimulatedOfficeSipsRunner(certificate);
            var dependencies = new WordMacroProjectSigningDependencies(
                runner, new ValidationSnapshotMutationMacroSigningPlatform());
            var options = new WordMacroProjectSigningOptions { OfficeSipsDirectory = toolsDirectory };
            options.Inspection.CmsVerification.CertificateValidation.ChainEvaluator = (_, _) => true;

            WordMacroProjectSigningResult result = WordMacroProjectSignatureService.TrySign(
                filePath, certificate.Thumbprint!, options, dependencies);

            Assert.False(result.Succeeded);
            Assert.False(result.MacroProjectPreserved);
            Assert.Contains(result.Findings, finding => finding.Code == "MacroValidatedSnapshotChanged");
        }

        private static void AddTimestampedMacroSignatureProfile(
            string filePath,
            WordMacroProjectSignatureProfile profile,
            X509Certificate2 certificate) {
            byte[] encoded = CreateAuthenticodeCms(profile, certificate);
            var signedData = new Org.BouncyCastle.Cms.CmsSignedData(encoded);
            Org.BouncyCastle.Cms.SignerInformation signer =
                signedData.GetSignerInfos().GetSigners().Single();
            Org.BouncyCastle.Asn1.DerObjectIdentifier timestampOid =
                Org.BouncyCastle.Asn1.Pkcs.PkcsObjectIdentifiers.IdAASignatureTimeStampToken;
            var timestampAttribute = new Org.BouncyCastle.Asn1.Cms.Attribute(
                timestampOid,
                new Org.BouncyCastle.Asn1.DerSet(new Org.BouncyCastle.Asn1.Asn1Encodable[] {
                    new Org.BouncyCastle.Asn1.DerSequence()
                }));
            var unsignedAttributes = new Org.BouncyCastle.Asn1.Cms.AttributeTable(
                new Dictionary<Org.BouncyCastle.Asn1.DerObjectIdentifier, object> {
                    [timestampOid] = timestampAttribute
                });
            Org.BouncyCastle.Cms.SignerInformation withTimestamp =
                Org.BouncyCastle.Cms.SignerInformation.ReplaceUnsignedAttributes(signer, unsignedAttributes);
            Org.BouncyCastle.Cms.CmsSignedData timestamped = Org.BouncyCastle.Cms.CmsSignedData.ReplaceSigners(
                signedData,
                new Org.BouncyCastle.Cms.SignerInformationStore(new[] { withTimestamp }));
            AddRawMacroSignatureProfile(filePath, profile,
                CreateMacroSignatureContainer(timestamped.GetEncoded()));
        }

        private sealed class StagingReplacementMacroSigningPlatform : IWordMacroProjectPlatform {
            private readonly SimulatedOfficeSipsRunner _runner;

            internal StagingReplacementMacroSigningPlatform(SimulatedOfficeSipsRunner runner) => _runner = runner;

            internal bool ReplacedStagingPath { get; private set; }
            public bool IsWindows => true;

            public bool TryGetSubjectInterfacePackage(string filePath, out Guid subjectGuid, out string detail) {
                subjectGuid = new Guid("6E64D5BD-CEB0-4B66-B4A0-15AC71775C48");
                detail = "simulated Microsoft Office SIP";
                return true;
            }

            public WordMacroProjectContentBindingResult ValidateContentBinding(
                string filePath,
                string digestAlgorithmOid,
                byte[] expectedDigest) {
                string stagingPath = _runner.Invocations.Last(invocation =>
                    invocation.Arguments.Count > 0 && invocation.Arguments[0] == "sign").Arguments.Last();
                string replacementPath = stagingPath + ".replacement";
                File.WriteAllBytes(replacementPath, new byte[] { 0x01, 0x02, 0x03 });
                File.Replace(replacementPath, stagingPath, null);
                ReplacedStagingPath = true;
                return new WordMacroProjectContentBindingResult(true, true,
                    "simulated Office SIP digest match before staged-path replacement");
            }
        }

        private sealed class ValidationSnapshotMutationMacroSigningPlatform : IWordMacroProjectPlatform {
            public bool IsWindows => true;

            public bool TryGetSubjectInterfacePackage(string filePath, out Guid subjectGuid, out string detail) {
                subjectGuid = new Guid("6E64D5BD-CEB0-4B66-B4A0-15AC71775C48");
                detail = "simulated Microsoft Office SIP";
                return true;
            }

            public WordMacroProjectContentBindingResult ValidateContentBinding(
                string filePath,
                string digestAlgorithmOid,
                byte[] expectedDigest) {
                File.WriteAllBytes(filePath, new byte[] { 0x01, 0x02, 0x03 });
                return new WordMacroProjectContentBindingResult(
                    true, true, "simulated binding before validation-snapshot mutation");
            }
        }
    }
}

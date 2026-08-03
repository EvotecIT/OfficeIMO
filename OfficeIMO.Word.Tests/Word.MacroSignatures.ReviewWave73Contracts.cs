using System.IO.Packaging;
using System.Security.Cryptography.X509Certificates;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void MacroSigningRejectsDuplicateRequiredProfileBeforeCommit() {
            string filePath = CreateMacroEnabledTestDocument("MacroSignatureDuplicateProfile.docm");
            byte[] originalBytes = File.ReadAllBytes(filePath);
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            string toolsDirectory = CreateFakeOfficeSipsDirectory();
            int profileIndex = 0;
            var runner = new RecordingMacroToolRunner(invocation => {
                if (invocation.Arguments.Count > 0 && invocation.Arguments[0] == "sign") {
                    profileIndex++;
                    WordMacroProjectSignatureProfile profile = (WordMacroProjectSignatureProfile)profileIndex;
                    string stagingPath = invocation.Arguments[invocation.Arguments.Count - 1];
                    AddMacroSignatureProfile(stagingPath, profile, certificate);
                    if (profile == WordMacroProjectSignatureProfile.V3) {
                        AddDuplicateLegacyMacroSignatureRelationship(stagingPath, certificate);
                    }
                }
                return Success();
            });
            var dependencies = new WordMacroProjectSigningDependencies(
                runner, new TestMacroSigningPlatform(isWindows: true));
            var options = new WordMacroProjectSigningOptions { OfficeSipsDirectory = toolsDirectory };
            TrustMacroTestCertificate(options.Inspection.CmsVerification.CertificateValidation);

            WordMacroProjectSigningResult result = WordMacroProjectSignatureService.TrySign(
                filePath, certificate.Thumbprint!, options, dependencies);

            Assert.False(result.Succeeded);
            Assert.True(result.MacroProjectPreserved);
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
            Assert.Contains(result.Findings, finding =>
                finding.Code == "DuplicateMacroSignatureProfile" &&
                finding.Profile == WordMacroProjectSignatureProfile.Legacy);
            Assert.Contains(result.Findings, finding =>
                finding.Code == "MacroSignatureProfilePolicyFailed" &&
                finding.Profile == WordMacroProjectSignatureProfile.Legacy);
        }

        private static void AddDuplicateLegacyMacroSignatureRelationship(
            string filePath,
            X509Certificate2 certificate) {
            Uri vbaUri;
            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, false)) {
                vbaUri = document.MainDocumentPart!.VbaProjectPart!.Uri;
            }

            byte[] cms = CreateAuthenticodeCms(WordMacroProjectSignatureProfile.Legacy, certificate);
            byte[] bytes = CreateMacroSignatureContainer(cms);
            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite);
            PackagePart vbaPart = package.GetPart(vbaUri);
            const string partName = "vbaProjectSignatureDuplicate.bin";
            Uri partUri = new Uri("/word/" + partName, UriKind.Relative);
            PackagePart signaturePart = package.CreatePart(
                partUri,
                "application/vnd.ms-office.vbaProjectSignature");
            using (Stream stream = signaturePart.GetStream(FileMode.Create, FileAccess.Write)) {
                stream.Write(bytes, 0, bytes.Length);
            }
            vbaPart.CreateRelationship(
                new Uri(partName, UriKind.Relative),
                TargetMode.Internal,
                WordMacroProjectSignatureInspector.LegacyRelationship);
        }
    }
}

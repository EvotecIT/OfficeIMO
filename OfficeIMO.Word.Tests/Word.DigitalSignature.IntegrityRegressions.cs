using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_DigitalSignature_EnvelopedTransformAuthenticatesManifestWhenTargetHasNoSignatureNode() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureNoOpEnvelopedManifest.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("No-op enveloped transform manifest");
                document.Save();
            }
            string documentDigest = ComputePackagePartSha256Digest(filePath, "/word/document.xml");
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\">" +
                "<SignedInfo><SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                "<Reference URI=\"#signed-object\"><Transforms><Transform Algorithm=\"http://www.w3.org/2000/09/xmldsig#enveloped-signature\" /></Transforms>" +
                "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>T2ZmaWNlSU1P</DigestValue></Reference></SignedInfo>" +
                "<Object Id=\"signed-object\"><Manifest><Reference URI=\"/word/document.xml\">" +
                "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>" + documentDigest + "</DigestValue>" +
                "</Reference></Manifest></Object></Signature>");
            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(new WordSignatureValidationOptions {
                ValidateCryptographicSignature = false
            });

            WordSignaturePartInfo part = Assert.Single(validation.SignatureInfo.SignatureParts);
            WordSignatureReferenceInfo packageReference = Assert.Single(
                part.SignedReferences,
                reference => reference.IsPackagePartReference);
            Assert.Equal(WordSignatureValidationState.Passed, packageReference.DigestVerificationStatus);
            Assert.Equal(WordSignatureValidationState.Passed, validation.SignedPartCoverageStatus);
            Assert.DoesNotContain(part.UnsupportedDetails, detail =>
                detail.Contains("does not preserve", System.StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Test_DigitalSignature_SignPackageRejectsAReplacedValidatedStage() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureReplacedStage.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Original package state");
                document.Save();
            }
            byte[] originalBytes = File.ReadAllBytes(filePath);

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            var options = new OfficePackageSigningOptions {
                BeforeCommit = (staging, _) => File.WriteAllBytes(staging, new byte[] { 0x01, 0x02, 0x03 })
            };

            OfficePackageSigningResult result = OfficePackageSignatureWriter.Sign(filePath, certificate, options);

            Assert.False(result.Succeeded);
            Assert.Contains(result.Details, detail =>
                detail.Contains("validated staging package changed", System.StringComparison.OrdinalIgnoreCase));
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
            using WordprocessingDocument preserved = WordprocessingDocument.Open(filePath, false);
            Assert.Null(preserved.DigitalSignatureOriginPart);
        }

        [Fact]
        public void Test_DigitalSignature_SigningBoundsRelationshipSelectorsBeforeSignatureXmlConstruction() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureRelationshipSelectorBudget.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Relationship selector budget");
                document.Save();
            }
            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = package.MainDocumentPart!;
                for (int index = 0; index < 128; index++) {
                    mainPart.AddExternalRelationship(
                        "urn:officeimo:relationship-budget",
                        new System.Uri("https://example.test/resource/" + index, System.UriKind.Absolute),
                        "rBudget" + index);
                }
            }
            byte[] originalBytes = File.ReadAllBytes(filePath);

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult bounded = WordDocument.TrySignPackage(
                filePath,
                certificate,
                new WordPackageSigningOptions { MaxSignedReferences = 32 });

            Assert.False(bounded.Succeeded);
            Assert.Contains(bounded.Details, detail =>
                detail.Contains("relationship selectors", System.StringComparison.OrdinalIgnoreCase));
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));

            WordPackageSigningResult permissive = WordDocument.TrySignPackage(
                filePath,
                certificate,
                new WordPackageSigningOptions { MaxSignedReferences = 512 });

            Assert.True(permissive.Succeeded, string.Join(System.Environment.NewLine, permissive.Details));
            Assert.True(permissive.SignedRelationshipSelectorCount >= 128);
        }
    }
}

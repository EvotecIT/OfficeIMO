using OfficeIMO.Word;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_DigitalSignature_DeduplicatesAdditionalCertificatesBeforeWriting() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureDuplicateCertificates.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Duplicate signing certificates");
                document.Save();
            }

            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            using X509Certificate2 additional = CreateSelfSignedSigningCertificate("CN=OfficeIMO Duplicate Additional Certificate");
            WordPackageSigningResult result = WordDocument.SignPackage(
                filePath,
                signer,
                new WordPackageSigningOptions {
                    AdditionalCertificates = Enumerable.Repeat(additional, 64).ToArray(),
                    MaxCertificates = 2
                });

            Assert.True(result.Succeeded);
            Assert.True(result.CreatedSignatureReadbackSucceeded);
            Assert.DoesNotContain(result.ValidationReport!.Diagnostics, diagnostic =>
                diagnostic.Code == "SignatureResourceLimitExceeded");
        }

        [Fact]
        public void Test_DigitalSignature_RejectsGeneratedSignatureOutsideConfiguredLimitAtomically() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureGeneratedXmlLimit.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Generated signature XML limit");
                document.Save();
            }
            byte[] originalBytes = File.ReadAllBytes(filePath);

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult result = WordDocument.TrySignPackage(
                filePath,
                certificate,
                new WordPackageSigningOptions { MaxSignatureBytes = 512 });

            Assert.False(result.Succeeded);
            Assert.Contains(result.Details, detail => detail.Contains("signature XML exceeds", StringComparison.OrdinalIgnoreCase));
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
        }

        [Fact]
        public void Test_DigitalSignature_BoundsAggregateLocalReferenceDigestWork() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureLocalReferenceDigestBudget.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Local SignedInfo reference work budget");
                document.Save();
            }
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            string encodedCertificate = Convert.ToBase64String(certificate.Export(X509ContentType.Cert));
            string digest = "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>T2ZmaWNlSU1P</DigestValue>";
            string reference = "<Reference URI=\"#payload\">" + digest + "</Reference>";
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\">" +
                "<SignedInfo><CanonicalizationMethod Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\" />" +
                "<SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                reference + reference + "</SignedInfo>" +
                "<SignatureValue>AA==</SignatureValue>" +
                "<KeyInfo><X509Data><X509Certificate>" + encodedCertificate + "</X509Certificate></X509Data></KeyInfo>" +
                "<Object Id=\"payload\">" + new string('x', 512) + "</Object></Signature>");
            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            WordSignatureValidationReport bounded = loaded.ValidateSignatures(new WordSignatureValidationOptions {
                MaxTotalDigestBytes = 768
            });
            WordSignatureValidationReport allowed = loaded.ValidateSignatures(new WordSignatureValidationOptions {
                MaxTotalDigestBytes = 4096
            });

            Assert.Contains(bounded.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
            Assert.DoesNotContain(allowed.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
            Assert.Contains(allowed.Diagnostics, finding => finding.Code == "XmlSignatureInvalid");
        }

        [Fact]
        public void Test_DigitalSignature_BoundsLocalReferenceTransformWork() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureLocalTransformBudget.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Local SignedInfo transform work budget");
                document.Save();
            }
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            string encodedCertificate = Convert.ToBase64String(certificate.Export(X509ContentType.Cert));
            string transform = "<Transform Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\" />";
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\">" +
                "<SignedInfo><CanonicalizationMethod Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\" />" +
                "<SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                "<Reference URI=\"#payload\"><Transforms>" + transform + transform + transform + "</Transforms>" +
                "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>T2ZmaWNlSU1P</DigestValue></Reference>" +
                "</SignedInfo><SignatureValue>AA==</SignatureValue>" +
                "<KeyInfo><X509Data><X509Certificate>" + encodedCertificate + "</X509Certificate></X509Data></KeyInfo>" +
                "<Object Id=\"payload\">" + new string('x', 512) + "</Object></Signature>");
            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(new WordSignatureValidationOptions {
                MaxTotalDigestBytes = 1024
            });

            Assert.Contains(validation.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
        }
    }
}

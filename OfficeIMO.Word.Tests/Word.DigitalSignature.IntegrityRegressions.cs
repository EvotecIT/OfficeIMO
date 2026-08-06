using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_DigitalSignature_CanonicalizationIgnoresConflictingXmlEncodingDeclaration() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureUtf16Canonicalization.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("UTF-16 canonicalization bridge");
                document.Save();
            }
            var xml = new XmlDocument { PreserveWhitespace = true };
            xml.LoadXml("<?xml version=\"1.0\" encoding=\"utf-16\"?><Root><Value>OfficeIMO</Value></Root>");
            using var archive = new OfficePackageSignatureArchive(
                File.ReadAllBytes(filePath),
                securityProvider: SecurityProvider);

            byte[] canonical = archive.Canonicalize(xml);

            Assert.Equal("<Root><Value>OfficeIMO</Value></Root>", Encoding.UTF8.GetString(canonical));
        }

        [Fact]
        public void Test_DigitalSignature_PackageReferenceRequiresContentTypeBinding() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureMissingContentType.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Unsigned content-type mapping");
                document.Save();
            }
            XNamespace ds = "http://www.w3.org/2000/09/xmldsig#";
            var reference = new XElement(ds + "Reference",
                new XAttribute("URI", "/word/document.xml"),
                new XElement(ds + "DigestMethod", new XAttribute("Algorithm", "http://www.w3.org/2001/04/xmlenc#sha256")),
                new XElement(ds + "DigestValue", ComputePackagePartSha256Digest(filePath, "/word/document.xml")));
            using var archive = new OfficePackageSignatureArchive(File.ReadAllBytes(filePath));

            OfficePackageDigestResult result = archive.VerifyReference(reference, 16 * 1024 * 1024);

            Assert.Equal(OfficeIMO.Security.OfficePackageSignatureValidationState.Failed, result.Status);
            Assert.Contains("not bound to an OPC content type", result.Detail, StringComparison.OrdinalIgnoreCase);
        }

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
                "<Object Id=\"signed-object\"><Manifest><Reference URI=\"/word/document.xml?ContentType=application%2Fvnd.openxmlformats-officedocument.wordprocessingml.document.main%2Bxml\">" +
                "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>" + documentDigest + "</DigestValue>" +
                "</Reference></Manifest></Object></Signature>");
            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly
            });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(SecurityProvider, new WordSignatureValidationOptions {
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

            OfficePackageSigningResult result = OfficePackageSignatureWriter.Sign(
                filePath,
                certificate,
                SecurityProvider,
                options);

            Assert.False(result.Succeeded);
            Assert.Contains(result.Details, detail =>
                detail.Contains("validated staging package changed", System.StringComparison.OrdinalIgnoreCase));
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
            using WordprocessingDocument preserved = WordprocessingDocument.Open(filePath, false);
            Assert.Null(preserved.DigitalSignatureOriginPart);
        }

        [Fact]
        public void Test_DigitalSignature_SigningRejectsFailedValidationReadbackBeforeCommit() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureFailedPreCommitReadback.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Original signed content");
                document.Save();
            }
            byte[] originalBytes = File.ReadAllBytes(filePath);

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            var options = new WordPackageSigningOptions {
                BeforeValidation = stagingPath => TamperDocumentText(stagingPath, "Tampered staged content")
            };

            WordPackageSigningResult result = WordDocument.TrySignPackage(filePath, SecurityProvider, certificate, options);

            Assert.False(result.Succeeded);
            Assert.NotNull(result.CreatedSignatureValidation);
            Assert.Equal(WordSignatureValidationState.Failed, result.CreatedSignatureValidation!.SignaturePart.SignedReferences
                .Single(reference => reference.Uri.StartsWith("/word/document.xml", StringComparison.OrdinalIgnoreCase))
                .DigestVerificationStatus);
            Assert.Contains(result.Details, detail =>
                detail.Contains("before atomic commit", StringComparison.OrdinalIgnoreCase));
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
            using WordprocessingDocument preserved = WordprocessingDocument.Open(filePath, false);
            Assert.Null(preserved.DigitalSignatureOriginPart);

            WordPackageSigningException exception = Assert.Throws<WordPackageSigningException>(() =>
                WordDocument.SignPackage(filePath, SecurityProvider, certificate, options));
            Assert.False(exception.Result.Succeeded);
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
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
            WordPackageSigningResult bounded = WordDocument.TrySignPackage(filePath, SecurityProvider,
                certificate,
                new WordPackageSigningOptions { MaxSignedReferences = 32 });

            Assert.False(bounded.Succeeded);
            Assert.Contains(bounded.Details, detail =>
                detail.Contains("relationship selectors", System.StringComparison.OrdinalIgnoreCase));
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));

            WordPackageSigningResult permissive = WordDocument.TrySignPackage(filePath, SecurityProvider,
                certificate,
                new WordPackageSigningOptions { MaxSignedReferences = 512 });

            Assert.True(permissive.Succeeded, string.Join(System.Environment.NewLine, permissive.Details));
            Assert.True(permissive.SignedRelationshipSelectorCount >= 128);
        }

        [Fact]
        public void Test_DigitalSignature_SigningAuthenticatesCustomSignatureLikeRelationshipTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureCustomRelationshipType.docx");
            const string relationshipId = "rCustomSignatureLike";
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Custom relationship type");
                document.Save();
            }
            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true)) {
                package.MainDocumentPart!.AddExternalRelationship(
                    "https://example.com/digital-signature/attachment",
                    new System.Uri("https://example.test/original", System.UriKind.Absolute),
                    relationshipId);
            }

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult signed = WordDocument.SignPackage(filePath, SecurityProvider, certificate);

            Assert.True(signed.Succeeded, string.Join(System.Environment.NewLine, signed.Details));
            XElement relationshipReference;
            using (var archive = ZipFile.OpenRead(filePath)) {
                ZipArchiveEntry signatureEntry = archive.Entries.Single(entry =>
                    entry.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase) &&
                    entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
                XDocument signature = XDocument.Load(signatureEntry.Open());
                relationshipReference = signature.Descendants().Single(element =>
                    element.Name.LocalName == "Reference" &&
                    element.Descendants().Any(descendant =>
                        descendant.Name.LocalName == "RelationshipReference" &&
                        string.Equals((string?)descendant.Attribute("SourceId"), relationshipId, StringComparison.Ordinal)));
            }

            RetargetRelationship(filePath, "word/_rels/document.xml.rels", relationshipId);
            using var tamperedArchive = new OfficePackageSignatureArchive(
                File.ReadAllBytes(filePath),
                securityProvider: SecurityProvider);
            OfficePackageDigestResult digest = tamperedArchive.VerifyReference(
                relationshipReference,
                maxPartBytes: 16 * 1024 * 1024);

            Assert.Equal(OfficeIMO.Security.OfficePackageSignatureValidationState.Failed, digest.Status);

            using WordDocument tampered = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly
            });
            var validationOptions = new WordSignatureValidationOptions();
            validationOptions.CertificateValidation.ChainEvaluator = static (_, _) => true;
            WordSignatureValidationReport validation = tampered.ValidateSignatures(SecurityProvider, validationOptions);
            Assert.False(validation.IsValidUnderPolicy);
            Assert.Equal(
                WordSignatureValidationState.Failed,
                validation.SignedPartDigestStatus);
            Assert.False(Assert.Single(validation.Signatures).IsValidUnderPolicy);
        }

        private static void RetargetRelationship(string filePath, string entryName, string relationshipId) {
            using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry entry = archive.GetEntry(entryName)!;
            XDocument relationships;
            using (Stream input = entry.Open()) relationships = XDocument.Load(input);
            XElement relationship = relationships.Descendants().Single(element =>
                element.Name.LocalName == "Relationship" &&
                string.Equals((string?)element.Attribute("Id"), relationshipId, StringComparison.Ordinal));
            relationship.SetAttributeValue("Target", "https://example.test/retargeted");
            entry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using Stream output = replacement.Open();
            relationships.Save(output);
        }

        private static void TamperDocumentText(string filePath, string replacementText) {
            using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            const string entryName = "word/document.xml";
            ZipArchiveEntry entry = archive.GetEntry(entryName)!;
            XDocument document;
            using (Stream input = entry.Open()) document = XDocument.Load(input);
            XElement text = document.Descendants().First(element => element.Name.LocalName == "t");
            text.Value = replacementText;
            entry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using Stream output = replacement.Open();
            document.Save(output);
        }
    }
}

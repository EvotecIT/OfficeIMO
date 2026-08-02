using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_DigitalSignature_ValidationSnapshotIncludesPendingBinaryFeedData() {
            string filePath = CreateSignedDocumentWithImagePart("WordDigitalSignaturePendingBinaryFeedData.docx");
            using WordDocument loaded = WordDocument.Load(filePath);
            ImagePart imagePart = Assert.Single(loaded._wordprocessingDocument.MainDocumentPart!.ImageParts);
            using (var replacement = new MemoryStream(new byte[] { 9, 8, 7, 6, 5 })) {
                imagePart.FeedData(replacement);
            }

            AssertLivePackageSignatureInvalid(loaded);
        }

        [Fact]
        public void Test_DigitalSignature_ValidationSnapshotIncludesPendingPartAddition() {
            string filePath = CreateSignedDocument("WordDigitalSignaturePendingPartAddition.docx");
            using WordDocument loaded = WordDocument.Load(filePath);
            ImagePart imagePart = loaded._wordprocessingDocument.MainDocumentPart!.AddImagePart(ImagePartType.Png);
            string relationshipId = loaded._wordprocessingDocument.MainDocumentPart.GetIdOfPart(imagePart);
            using (var content = new MemoryStream(new byte[] { 1, 3, 5, 7, 9 })) {
                imagePart.FeedData(content);
            }

            byte[] snapshot = (byte[])typeof(WordDocument)
                .GetMethod("CreateSignatureValidationSnapshot", BindingFlags.Instance | BindingFlags.NonPublic)!
                .Invoke(loaded, new object[] { new WordSignatureValidationOptions() })!;
            using var archive = new ZipArchive(new MemoryStream(snapshot), ZipArchiveMode.Read);
            Assert.NotNull(archive.GetEntry(imagePart.Uri.ToString().TrimStart('/')));
            using Stream relationships = archive.GetEntry("word/_rels/document.xml.rels")!.Open();
            XDocument relationshipDocument = XDocument.Load(relationships);
            Assert.Contains(relationshipDocument.Root!.Elements(), element =>
                string.Equals(element.Attribute("Id")?.Value, relationshipId, System.StringComparison.Ordinal));
        }

        [Fact]
        public void Test_DigitalSignature_ValidationSnapshotIncludesPendingPartRemoval() {
            string filePath = CreateSignedDocumentWithImagePart("WordDigitalSignaturePendingPartRemoval.docx");
            using WordDocument loaded = WordDocument.Load(filePath);
            MainDocumentPart mainPart = loaded._wordprocessingDocument.MainDocumentPart!;
            mainPart.DeletePart(Assert.Single(mainPart.ImageParts));

            AssertLivePackageSignatureInvalid(loaded);
        }

        [Fact]
        public void Test_DigitalSignature_PartLimitIsIndeterminateRatherThanFailedDigest() {
            string filePath = CreateSignedDocument("WordDigitalSignaturePartLimitStatus.docx");
            const string partUri = "/word/document.xml";
            string digest = ComputePackagePartSha256Digest(filePath, partUri);
            XNamespace ds = "http://www.w3.org/2000/09/xmldsig#";
            var reference = new XElement(ds + "Reference",
                new XAttribute(
                    "URI",
                    partUri + "?ContentType=application%2Fvnd.openxmlformats-officedocument.wordprocessingml.document.main%2Bxml"),
                new XElement(ds + "DigestMethod",
                    new XAttribute("Algorithm", "http://www.w3.org/2001/04/xmlenc#sha256")),
                new XElement(ds + "DigestValue", digest));
            using var archive = new OfficePackageSignatureArchive(File.ReadAllBytes(filePath));

            OfficePackageDigestResult result = archive.VerifyReference(reference, maxPartBytes: 16);

            Assert.Equal(OfficePackageSignatureDigestVerificationStatus.Unsupported, result.Status);
            Assert.Contains("byte limit", result.Detail, System.StringComparison.OrdinalIgnoreCase);
        }

        private string CreateSignedDocument(string fileName) {
            string filePath = Path.Combine(_directoryWithFiles, fileName);
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed validation snapshot");
                document.Save();
            }
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult signing = WordDocument.SignPackage(filePath, certificate);
            Assert.True(signing.Succeeded, string.Join(System.Environment.NewLine, signing.Details));
            return filePath;
        }

        private string CreateSignedDocumentWithImagePart(string fileName) {
            string filePath = Path.Combine(_directoryWithFiles, fileName);
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed validation snapshot with binary part");
                document.Save();
            }
            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true)) {
                ImagePart imagePart = package.MainDocumentPart!.AddImagePart(ImagePartType.Png);
                using var content = new MemoryStream(new byte[] { 1, 2, 3, 4, 5 });
                imagePart.FeedData(content);
            }
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult signing = WordDocument.SignPackage(filePath, certificate);
            Assert.True(signing.Succeeded, string.Join(System.Environment.NewLine, signing.Details));
            return filePath;
        }

        private static void AssertLivePackageSignatureInvalid(WordDocument document) {
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;
            WordSignatureValidationReport validation = document.ValidateSignatures(options);

            Assert.NotEqual(WordSignatureValidationState.Passed, validation.SignedPartDigestStatus);
            Assert.False(validation.IsValidUnderPolicy);
            Assert.False(Assert.Single(validation.Signatures).IsValidUnderPolicy);
        }
    }
}

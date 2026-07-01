using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Collections.Generic;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Shared;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_DigitalSignature_MissingPart_ReturnsNull() {
            string tempFile = Path.GetTempFileName();
            using (WordDocument document = WordDocument.Create(tempFile)) {
                Assert.True(document.ApplicationProperties.DigitalSignature == null);
                WordSignatureInfo signatures = document.InspectSignatures();
                Assert.False(signatures.HasSignatures);
                Assert.Equal(0, signatures.FindingCount);

                WordSignatureValidationReport validation = document.ValidateSignatures();
                Assert.False(validation.HasSignatures);
                Assert.Equal(WordSignatureValidationState.NotPresent, validation.PackageStructureStatus);
                Assert.Equal(WordSignatureValidationState.NotPresent, validation.XmlSignatureStatus);
                Assert.Equal(WordSignatureValidationState.NotPresent, validation.CryptographicStatus);
            }
        }

        [Fact]
        public void Test_DigitalSignature_PartDeleted_ReturnsNull() {
            string tempFile = Path.GetTempFileName();
            using (WordDocument document = WordDocument.Create(tempFile)) {
                document.ApplicationProperties.DigitalSignature = new DigitalSignature();
                Assert.True(document.ApplicationProperties.DigitalSignature != null);
                var extendedPart = document._wordprocessingDocument!.ExtendedFilePropertiesPart;
                Assert.NotNull(extendedPart);
                document._wordprocessingDocument!.DeletePart(extendedPart);
                Assert.True(document.ApplicationProperties.DigitalSignature == null);
            }
        }

        [Fact]
        public void Test_DigitalSignature_InspectSignaturesReportsPackageMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureMetadata.docx");
            byte[] signatureBytes = CreateSignatureXml();

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed metadata carrier");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordSignatureInfo signatures = document.InspectSignatures();

                Assert.True(signatures.HasSignatures);
                Assert.True(signatures.HasDigitalSignatureOriginPart);
                Assert.True(signatures.HasApplicationSignatureMetadata);
                Assert.Contains("origin.sigs", signatures.OriginPartUri, System.StringComparison.OrdinalIgnoreCase);
                WordSignaturePartInfo signaturePart = Assert.Single(signatures.SignatureParts);
                Assert.Contains("_xmlsignatures", signaturePart.Uri, System.StringComparison.OrdinalIgnoreCase);
                Assert.Equal("http://www.w3.org/2001/04/xmldsig-more#rsa-sha256", signaturePart.SignatureMethodAlgorithm);
                Assert.Contains("http://www.w3.org/2001/04/xmlenc#sha256", signaturePart.DigestMethodAlgorithms);
                WordSignatureReferenceInfo signedReference = Assert.Single(signaturePart.SignedReferences);
                Assert.Equal("/word/document.xml", signedReference.Uri);
                Assert.Equal("http://www.w3.org/2001/04/xmlenc#sha256", signedReference.DigestMethodAlgorithm);
                Assert.True(signedReference.HasDigestValue);
                Assert.Equal("T2ZmaWNlSU1P", signedReference.DigestValue);
                Assert.True(signedReference.IsPackagePartReference);
                Assert.Equal("/word/document.xml", signedReference.TargetPartUri);
                Assert.True(signedReference.TargetPartExists);
                Assert.Contains("CN=OfficeIMO Test", signaturePart.X509SubjectNames);
                Assert.Contains(signatures.UnsupportedDetails, detail => detail.Contains("metadata-only", System.StringComparison.OrdinalIgnoreCase));

                WordSignatureValidationReport validation = document.ValidateSignatures();
                Assert.False(validation.IsStructurallyValid);
                Assert.Equal(WordSignatureValidationState.Passed, validation.PackageStructureStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.XmlSignatureStatus);
                Assert.Equal(WordSignatureValidationState.NotChecked, validation.CryptographicStatus);
                Assert.Equal(WordSignatureValidationState.NotChecked, validation.CertificateChainStatus);
                Assert.Equal(WordSignatureValidationState.NotChecked, validation.RevocationStatus);
                Assert.Equal(WordSignatureValidationState.NotPresent, validation.TimestampStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.SignedPartCoverageStatus);
                Assert.Equal(WordSignatureValidationState.Failed, validation.SignedPartDigestStatus);
                Assert.Contains(validation.Findings, finding => finding.Contains("digest did not match", System.StringComparison.OrdinalIgnoreCase));
                Assert.Contains(validation.Findings, finding => finding.Contains("Cryptographic signature validation is not performed", System.StringComparison.OrdinalIgnoreCase));
                Assert.Contains(validation.Findings, finding => finding.Contains("package-part references resolve", System.StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_DigitalSignature_InspectSignaturesReportsTimestampMetadataWithoutValidationClaim() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureTimestampMetadata.docx");
            const string opcTimestampValue = "2026-06-30T08:15:30Z";
            const string xadesTimestampValue = "2026-06-30T08:16:30Z";

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Timestamp metadata carrier");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(
                filePath,
                CreateSignatureXml(
                    includeOpcSignatureTime: true,
                    opcSignatureTimeValue: opcTimestampValue,
                    includeXadesSigningTime: true,
                    xadesSigningTimeValue: xadesTimestampValue));

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordSignatureValidationReport validation = document.ValidateSignatures();
                WordSignaturePartInfo signaturePart = Assert.Single(validation.SignatureInfo.SignatureParts);

                Assert.Equal(2, signaturePart.Timestamps.Count);
                Assert.Contains(signaturePart.Timestamps, timestamp =>
                    timestamp.Kind == "OPC SignatureTime" &&
                    timestamp.Value == opcTimestampValue &&
                    timestamp.Format == "YYYY-MM-DDThh:mm:ssTZD");
                Assert.Contains(signaturePart.Timestamps, timestamp =>
                    timestamp.Kind == "XAdES SigningTime" &&
                    timestamp.Value == xadesTimestampValue &&
                    timestamp.Format == null);
                Assert.Equal(WordSignatureValidationState.NotChecked, validation.TimestampStatus);
                Assert.Contains(validation.Findings, finding => finding.Contains("timestamp metadata was found", System.StringComparison.OrdinalIgnoreCase));
                Assert.Contains(validation.SignatureInfo.Details, detail => detail.Contains("Signature timestamp", System.StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesSupportsSignedFixture() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "DigitalSignatures", "signed-valid.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing signed DOCX fixture: {sourcePath}");

            using (WordDocument document = WordDocument.Load(sourcePath, readOnly: true)) {
                WordSignatureInfo signatures = document.InspectSignatures();

                Assert.True(signatures.HasSignatures);
                Assert.True(signatures.HasDigitalSignatureOriginPart);
                Assert.NotEmpty(signatures.SignatureParts);
                Assert.Contains(signatures.SignatureParts.SelectMany(part => part.X509SubjectNames), subject =>
                    subject.Contains("OfficeIMO Fixture Package Signing", System.StringComparison.OrdinalIgnoreCase));

                WordSignatureValidationReport validation = document.ValidateSignatures();

                Assert.True(validation.IsStructurallyValid);
                Assert.Equal(WordSignatureValidationState.Passed, validation.PackageStructureStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.XmlSignatureStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.SignedPartCoverageStatus);
                Assert.NotEqual(WordSignatureValidationState.Failed, validation.SignedPartDigestStatus);
                Assert.Equal(WordSignatureValidationState.NotChecked, validation.CryptographicStatus);
                Assert.Equal(WordSignatureValidationState.NotChecked, validation.CertificateChainStatus);
                Assert.Equal(WordSignatureValidationState.NotChecked, validation.RevocationStatus);
                Assert.Equal(WordSignatureValidationState.NotChecked, validation.TimestampStatus);
                Assert.Contains(validation.SignatureInfo.SignatureParts.SelectMany(part => part.Timestamps), timestamp =>
                    !string.IsNullOrWhiteSpace(timestamp.Value));
                Assert.Contains(validation.SignatureInfo.SignatureParts.SelectMany(part => part.SignedReferences), reference =>
                    reference.HasDigestValue &&
                    reference.IsPackagePartReference &&
                    reference.TargetPartExists == true);
            }
        }

        [Fact]
        public void Test_DigitalSignature_SharedPackageInspectorReadsOpenXmlSignatureMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureSharedInspector.docx");
            byte[] signatureBytes = CreateSignatureXml();

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Shared package signature inspector carrier");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, false)) {
                OfficePackageSignatureInfo signatures = OfficePackageSignatureInspector.Inspect(
                    package,
                    package.DigitalSignatureOriginPart,
                    package.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null);

                Assert.True(signatures.HasSignatures);
                Assert.True(signatures.HasDigitalSignatureOriginPart);
                Assert.True(signatures.HasApplicationSignatureMetadata);
                OfficePackageSignaturePartInfo signaturePart = Assert.Single(signatures.SignatureParts);
                OfficePackageSignatureReferenceInfo signedReference = Assert.Single(signaturePart.SignedReferences);
                Assert.Equal("/word/document.xml", signedReference.TargetPartUri);
                Assert.True(signedReference.TargetPartExists);
                Assert.True(signedReference.HasDigestValue);
                Assert.Equal(OfficePackageSignatureDigestVerificationStatus.Failed, signedReference.DigestVerificationStatus);
                Assert.Contains(signatures.Details, detail => detail.Contains("Signed reference", System.StringComparison.OrdinalIgnoreCase));
                Assert.Contains(signatures.UnsupportedDetails, detail => detail.Contains("metadata-only", System.StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesVerifiesSimplePackagePartDigest() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureValidSimpleDigest.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Simple digest verification carrier");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(filePath, CreateSignatureXml(digestValue: ComputePackagePartSha256Digest(filePath, "/word/document.xml")));

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordSignatureValidationReport validation = document.ValidateSignatures();

                WordSignatureReferenceInfo signedReference = Assert.Single(Assert.Single(validation.SignatureInfo.SignatureParts).SignedReferences);
                Assert.Equal(WordSignatureValidationState.Passed, signedReference.DigestVerificationStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.SignedPartDigestStatus);
                Assert.Contains(validation.Findings, finding => finding.Contains("digests match", System.StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesReportsMismatchedSimplePackagePartDigest() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureMismatchedSimpleDigest.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Mismatched digest verification carrier");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(filePath, CreateSignatureXml(digestValue: "T2ZmaWNlSU1P"));

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordSignatureValidationReport validation = document.ValidateSignatures();

                WordSignatureReferenceInfo signedReference = Assert.Single(Assert.Single(validation.SignatureInfo.SignatureParts).SignedReferences);
                Assert.Equal(WordSignatureValidationState.Failed, signedReference.DigestVerificationStatus);
                Assert.Equal(WordSignatureValidationState.Failed, validation.SignedPartDigestStatus);
                Assert.Contains(validation.Findings, finding => finding.Contains("digest did not match", System.StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesLeavesTransformedDigestVerificationUnsupported() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureTransformedDigestUnsupported.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Transformed digest verification carrier");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(
                filePath,
                CreateSignatureXml(
                    digestValue: ComputePackagePartSha256Digest(filePath, "/word/document.xml"),
                    transformAlgorithm: "http://www.w3.org/2000/09/xmldsig#enveloped-signature"));

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordSignatureValidationReport validation = document.ValidateSignatures();

                WordSignatureReferenceInfo signedReference = Assert.Single(Assert.Single(validation.SignatureInfo.SignatureParts).SignedReferences);
                Assert.Equal(WordSignatureValidationState.Unsupported, signedReference.DigestVerificationStatus);
                Assert.Equal("http://www.w3.org/2000/09/xmldsig#enveloped-signature", Assert.Single(signedReference.TransformAlgorithms));
                Assert.Equal(WordSignatureValidationState.Unsupported, validation.SignedPartDigestStatus);
                Assert.Contains(validation.Findings, finding => finding.Contains("transforms", System.StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesReportsMissingReferenceDigestValue() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureMissingDigestValue.docx");
            byte[] signatureBytes = CreateSignatureXml(includeDigestValue: false);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Missing reference digest value carrier");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordSignatureValidationReport validation = document.ValidateSignatures();

                WordSignatureReferenceInfo signedReference = Assert.Single(Assert.Single(validation.SignatureInfo.SignatureParts).SignedReferences);
                Assert.False(signedReference.HasDigestValue);
                Assert.Null(signedReference.DigestValue);
                Assert.False(validation.IsStructurallyValid);
                Assert.Equal(WordSignatureValidationState.Passed, validation.PackageStructureStatus);
                Assert.Equal(WordSignatureValidationState.Unsupported, validation.XmlSignatureStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.SignedPartCoverageStatus);
                Assert.Contains(validation.Findings, finding => finding.Contains("Reference DigestValue", System.StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesReportsMissingSignedPackagePartReference() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureMissingSignedPart.docx");
            byte[] signatureBytes = CreateSignatureXml("/word/missing.xml");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Missing signed package part carrier");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordSignatureValidationReport validation = document.ValidateSignatures();

                WordSignatureReferenceInfo signedReference = Assert.Single(Assert.Single(validation.SignatureInfo.SignatureParts).SignedReferences);
                Assert.True(signedReference.IsPackagePartReference);
                Assert.Equal("/word/missing.xml", signedReference.TargetPartUri);
                Assert.False(signedReference.TargetPartExists);
                Assert.False(validation.IsStructurallyValid);
                Assert.Equal(WordSignatureValidationState.Passed, validation.PackageStructureStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.XmlSignatureStatus);
                Assert.Equal(WordSignatureValidationState.Failed, validation.SignedPartCoverageStatus);
                Assert.Contains(validation.Findings, finding => finding.Contains("missing package part", System.StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesReportsMalformedXmlSignaturePart() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureMalformed.docx");
            byte[] signatureBytes = Encoding.UTF8.GetBytes("<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo>");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Malformed signature metadata carrier");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordSignatureValidationReport validation = document.ValidateSignatures();

                Assert.False(validation.IsStructurallyValid);
                Assert.Equal(WordSignatureValidationState.Passed, validation.PackageStructureStatus);
                Assert.Equal(WordSignatureValidationState.Failed, validation.XmlSignatureStatus);
                Assert.Contains(validation.Findings, finding => finding.Contains("could not be parsed", System.StringComparison.OrdinalIgnoreCase));
                Assert.True(Assert.Single(validation.SignatureInfo.SignatureParts).HasParseError);
            }
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesReportsApplicationMetadataWithoutOriginAsUnsupported() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureApplicationOnly.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Application signature metadata only");
                document.ApplicationProperties.DigitalSignature = new DigitalSignature();
                document.Save(false, new WordSaveOptions { SignedDocumentPolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation });
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordSignatureValidationReport validation = document.ValidateSignatures();

                Assert.True(validation.HasSignatures);
                Assert.False(validation.IsStructurallyValid);
                Assert.Equal(WordSignatureValidationState.Unsupported, validation.PackageStructureStatus);
                Assert.Equal(WordSignatureValidationState.NotPresent, validation.XmlSignatureStatus);
                Assert.Equal(WordSignatureValidationState.NotChecked, validation.CryptographicStatus);
                Assert.Contains(validation.Findings, finding => finding.Contains("no digital-signature origin part", System.StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_DigitalSignature_NoOpSavePreservesSignatureMetadataParts() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignaturePreserve.docx");
            byte[] signatureBytes = CreateSignatureXml();

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed no-op save");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.Save(false, new WordSaveOptions { SignedDocumentPolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation });
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, false)) {
                Assert.NotNull(package.DigitalSignatureOriginPart);
                XmlSignaturePart signaturePart = Assert.Single(package.DigitalSignatureOriginPart!.XmlSignatureParts);
                using Stream stream = signaturePart.GetStream(FileMode.Open, FileAccess.Read);
                using var buffer = new MemoryStream();
                stream.CopyTo(buffer);
                Assert.Equal(signatureBytes, buffer.ToArray());
                Assert.NotNull(package.ExtendedFilePropertiesPart?.Properties?.DigitalSignature);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordFeatureFinding signatures = Assert.Single(document.InspectFeatures().FindFeatures("Digital signatures"));

                Assert.Equal(WordFeatureSupportLevel.Unsupported, signatures.SupportLevel);
                Assert.Contains(signatures.Details, detail => detail.Contains("origin.sigs", System.StringComparison.OrdinalIgnoreCase));
                Assert.Contains(signatures.Details, detail => detail.Contains("_xmlsignatures", System.StringComparison.OrdinalIgnoreCase));
                Assert.Contains(signatures.Details, detail => detail.Contains("not validated", System.StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_DigitalSignature_SaveBlocksSignedDocumentByDefault() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureSaveBlocked.docx");
            byte[] signatureBytes = CreateSignatureXml();

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed blocked save");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AddParagraph("Mutation after signing");

                WordSignatureSavePolicyException exception = Assert.Throws<WordSignatureSavePolicyException>(() => document.Save(false));

                Assert.Equal("Save", exception.Operation);
                Assert.True(exception.SignatureInfo.HasSignatures);
                Assert.Contains("may invalidate existing signatures", exception.Message, System.StringComparison.OrdinalIgnoreCase);
                Assert.Contains("AllowSignatureInvalidation", exception.Message, System.StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_DigitalSignature_SaveAllowsExplicitInvalidationPolicy() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureSaveAllowed.docx");
            byte[] signatureBytes = CreateSignatureXml();

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed allowed save");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AddParagraph("Mutation after signing");
                document.Save(false, new WordSaveOptions { SignedDocumentPolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation });
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                Assert.Contains(document.Paragraphs, paragraph => paragraph.Text == "Mutation after signing");
                Assert.True(document.InspectSignatures().HasSignatures);
            }
        }

        [Fact]
        public void Test_DigitalSignature_SaveAsMemoryStreamBlocksSignedDocumentByDefault() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureStreamBlocked.docx");
            byte[] signatureBytes = CreateSignatureXml();

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed stream blocked");
                document.Save(false);
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Throws<WordSignatureSavePolicyException>(() => document.SaveAsMemoryStream());
                using MemoryStream stream = document.SaveAsMemoryStream(new WordSaveOptions { SignedDocumentPolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation });

                Assert.True(stream.Length > 0);
            }
        }

#if NET472
        [Fact]
        public void Test_DigitalSignature_SignPackageCreatesStructurallyReadableSignatureOnSupportedAdapter() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureSignedByAdapter.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Package signing adapter proof");
                document.Save(false);
            }

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult result = WordDocument.SignPackage(
                filePath,
                certificate,
                new WordPackageSigningOptions { SignatureId = "OfficeIMOTestSignature" });

            Assert.True(result.IsSupported);
            Assert.True(result.Succeeded);
            Assert.True(result.SignedPartCount > 0);
            Assert.True(result.SignedRelationshipSelectorCount > 0);
            Assert.True(result.SignatureCount > 0);
            Assert.Contains("package/services/digital-signature", result.SignaturePartUri, System.StringComparison.OrdinalIgnoreCase);
            Assert.NotNull(result.ValidationReport);
            Assert.True(result.ValidationReport!.IsStructurallyValid);
            Assert.Equal(WordSignatureValidationState.NotChecked, result.ValidationReport.CryptographicStatus);

            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, false)) {
                Assert.NotNull(package.DigitalSignatureOriginPart);
                Assert.NotEmpty(package.DigitalSignatureOriginPart!.XmlSignatureParts);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordSignatureValidationReport validation = document.ValidateSignatures();

                Assert.True(validation.HasSignatures);
                Assert.True(validation.IsStructurallyValid);
                Assert.Equal(WordSignatureValidationState.Passed, validation.PackageStructureStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.XmlSignatureStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.SignedPartCoverageStatus);
                Assert.Equal(WordSignatureValidationState.NotChecked, validation.CertificateChainStatus);
                Assert.True(validation.SignatureInfo.SignatureParts.Count > 0);
                Assert.Contains(validation.SignatureInfo.SignatureParts.SelectMany(part => part.SignedReferences), reference => reference.HasDigestValue);
            }
        }

        [Fact]
        public void Test_DigitalSignature_TrySignPackageFailsWhenRequestedPartIsMissing() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureMissingRequestedPart.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Package signing missing requested part proof");
                document.Save(false);
            }

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult result = WordDocument.TrySignPackage(
                filePath,
                certificate,
                new WordPackageSigningOptions {
                    PartUris = new[] { "/word/document.xml", "/word/missing-part.xml" },
                    SignatureId = "OfficeIMOMissingPartSignature"
                });

            Assert.True(result.IsSupported);
            Assert.False(result.Succeeded);
            Assert.Equal(0, result.SignedPartCount);
            Assert.Null(result.ValidationReport);
            Assert.Contains(result.Details, detail => detail.Contains("/word/missing-part.xml", System.StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Test_DigitalSignature_SelectiveSigningScopesPartRelationshipSelectors() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureSelectivePartRelationships.docx");
            string imagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Package signing selective relationship proof");
                document.Save(false);
            }

            int documentPartRelationshipCount;
            int headerPartRelationshipCount;
            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = package.MainDocumentPart!;
                HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
                ImagePart imagePart = headerPart.AddImagePart(ImagePartType.Png);
                using (FileStream stream = File.OpenRead(imagePath)) {
                    imagePart.FeedData(stream);
                }

                headerPart.Header = new Header(new Paragraph(new Run(new Text("Header image relationship carrier"))));
                string headerRelationshipId = mainPart.GetIdOfPart(headerPart);
                Body body = mainPart.Document.Body!;
                SectionProperties sectionProperties = body.Elements<SectionProperties>().LastOrDefault()
                    ?? body.AppendChild(new SectionProperties());
                sectionProperties.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerRelationshipId });
                mainPart.Document.Save();

                documentPartRelationshipCount = mainPart.Parts.Count();
                headerPartRelationshipCount = headerPart.Parts.Count();
            }

            Assert.True(headerPartRelationshipCount > 0);

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult result = WordDocument.SignPackage(
                filePath,
                certificate,
                new WordPackageSigningOptions {
                    IncludePackageRelationships = false,
                    IncludePartRelationships = true,
                    PartUris = new[] { "/word/document.xml" },
                    SignatureId = "OfficeIMOSelectivePartSignature"
                });

            Assert.True(result.IsSupported);
            Assert.True(result.Succeeded);
            Assert.Equal(1, result.SignedPartCount);
            Assert.Equal(documentPartRelationshipCount, result.SignedRelationshipSelectorCount);
            Assert.True(result.SignedRelationshipSelectorCount < documentPartRelationshipCount + headerPartRelationshipCount);
            Assert.NotNull(result.ValidationReport);
        }

        [Fact]
        public void Test_DigitalSignature_SignPackageCanResolveCertificateFromStoreOnSupportedAdapter() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureSignedByStoreCertificate.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Package signing certificate-store proof");
                document.Save(false);
            }

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            AddCertificateToCurrentUserStore(certificate);
            try {
                WordPackageSigningResult result = WordDocument.SignPackage(
                    filePath,
                    certificate.Thumbprint!,
                    new WordPackageCertificateStoreOptions {
                        StoreLocation = StoreLocation.CurrentUser,
                        StoreName = StoreName.My,
                        RequirePrivateKey = true,
                        IncludeInvalidCertificates = true
                    },
                    new WordPackageSigningOptions { SignatureId = "OfficeIMOStoreCertificateSignature" });

                Assert.True(result.IsSupported);
                Assert.True(result.Succeeded);
                Assert.True(result.SignedPartCount > 0);
                Assert.NotNull(result.ValidationReport);
                Assert.True(result.ValidationReport!.IsStructurallyValid);
                Assert.Contains(result.ValidationReport.SignatureInfo.SignatureParts, part =>
                    part.SignedReferences.Any(reference => reference.HasDigestValue));
            } finally {
                RemoveCertificateFromCurrentUserStore(certificate.Thumbprint);
            }
        }
#else
        [Fact]
        public void Test_DigitalSignature_TrySignPackageReportsUnsupportedAdapter() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureUnsupportedSigningAdapter.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Unsupported package signing adapter proof");
                document.Save(false);
            }

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult result = WordDocument.TrySignPackage(filePath, certificate);

            Assert.False(result.IsSupported);
            Assert.False(result.Succeeded);
            Assert.Null(result.ValidationReport);
            Assert.Contains(result.Details, detail => detail.Contains("PackageDigitalSignatureManager", System.StringComparison.OrdinalIgnoreCase));

            WordPackageSigningException exception = Assert.Throws<WordPackageSigningException>(() => WordDocument.SignPackage(filePath, certificate));
            Assert.False(exception.Result.IsSupported);
            Assert.False(exception.Result.Succeeded);
        }
#endif

        [Fact]
        public void Test_DigitalSignature_TrySignPackageReportsMissingStoreCertificate() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureMissingStoreCertificate.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Missing store certificate proof");
                document.Save(false);
            }

            WordPackageSigningResult result = WordDocument.TrySignPackage(
                filePath,
                "00 11 22 33 44 55 66 77 88 99 AA BB CC DD EE FF 00 11 22 33",
                new WordPackageCertificateStoreOptions {
                    StoreLocation = StoreLocation.CurrentUser,
                    StoreName = StoreName.My
                });

            Assert.False(result.Succeeded);
            Assert.Null(result.ValidationReport);
            Assert.Contains(result.Details, detail => detail.Contains("was not found", System.StringComparison.OrdinalIgnoreCase));

            WordPackageSigningException exception = Assert.Throws<WordPackageSigningException>(() => WordDocument.SignPackage(
                filePath,
                "00112233445566778899AABBCCDDEEFF00112233"));
            Assert.False(exception.Result.Succeeded);
        }

        private static byte[] CreateSignatureXml(
            string referenceUri = "/word/document.xml",
            bool includeDigestValue = true,
            string? digestValue = null,
            string? transformAlgorithm = null,
            bool includeOpcSignatureTime = false,
            string? opcSignatureTimeValue = null,
            bool includeXadesSigningTime = false,
            string? xadesSigningTimeValue = null) {
            return Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\">" +
                "<SignedInfo>" +
                "<SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                "<Reference URI=\"" + referenceUri + "\">" +
                (string.IsNullOrWhiteSpace(transformAlgorithm) ? string.Empty : "<Transforms><Transform Algorithm=\"" + transformAlgorithm + "\" /></Transforms>") +
                "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" />" +
                (includeDigestValue ? "<DigestValue>" + (digestValue ?? "T2ZmaWNlSU1P") + "</DigestValue>" : string.Empty) +
                "</Reference>" +
                "</SignedInfo>" +
                "<KeyInfo><X509Data><X509SubjectName>CN=OfficeIMO Test</X509SubjectName></X509Data></KeyInfo>" +
                CreateSignatureTimestampXml(includeOpcSignatureTime, opcSignatureTimeValue, includeXadesSigningTime, xadesSigningTimeValue) +
                "</Signature>");
        }

        private static string CreateSignatureTimestampXml(
            bool includeOpcSignatureTime,
            string? opcSignatureTimeValue,
            bool includeXadesSigningTime,
            string? xadesSigningTimeValue) {
            var builder = new StringBuilder();
            if (includeOpcSignatureTime) {
                builder.Append("<Object><SignatureProperties><SignatureProperty Target=\"#OfficeIMOTestSignature\">");
                builder.Append("<mdssi:SignatureTime xmlns:mdssi=\"http://schemas.openxmlformats.org/package/2006/digital-signature\">");
                builder.Append("<mdssi:Format>YYYY-MM-DDThh:mm:ssTZD</mdssi:Format>");
                builder.Append("<mdssi:Value>");
                builder.Append(opcSignatureTimeValue ?? "2026-06-30T08:15:30Z");
                builder.Append("</mdssi:Value>");
                builder.Append("</mdssi:SignatureTime>");
                builder.Append("</SignatureProperty></SignatureProperties></Object>");
            }

            if (includeXadesSigningTime) {
                builder.Append("<Object><xades:QualifyingProperties xmlns:xades=\"http://uri.etsi.org/01903/v1.3.2#\">");
                builder.Append("<xades:SignedProperties><xades:SignedSignatureProperties><xades:SigningTime>");
                builder.Append(xadesSigningTimeValue ?? "2026-06-30T08:16:30Z");
                builder.Append("</xades:SigningTime></xades:SignedSignatureProperties></xades:SignedProperties>");
                builder.Append("</xades:QualifyingProperties></Object>");
            }

            return builder.ToString();
        }

        private static string ComputePackagePartSha256Digest(string filePath, string partUri) {
            using WordprocessingDocument package = WordprocessingDocument.Open(filePath, false);
            OpenXmlPart part = package.Parts
                .Select(pair => pair.OpenXmlPart)
                .SelectMany(EnumerateParts)
                .First(part => part.Uri.ToString().Equals(partUri, System.StringComparison.OrdinalIgnoreCase));

            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using SHA256 sha256 = SHA256.Create();
            return System.Convert.ToBase64String(sha256.ComputeHash(stream));
        }

        private static IEnumerable<OpenXmlPart> EnumerateParts(OpenXmlPart part) {
            yield return part;

            foreach (IdPartPair child in part.Parts) {
                foreach (OpenXmlPart descendant in EnumerateParts(child.OpenXmlPart)) {
                    yield return descendant;
                }
            }
        }

        private static void AddDigitalSignatureMetadata(string filePath, byte[] signatureBytes) {
            using WordprocessingDocument package = WordprocessingDocument.Open(filePath, true);
            package.AddDigitalSignatureOriginPart();
            XmlSignaturePart signaturePart = package.DigitalSignatureOriginPart!.AddNewPart<XmlSignaturePart>();
            using (var stream = new MemoryStream(signatureBytes)) {
                signaturePart.FeedData(stream);
            }

            ExtendedFilePropertiesPart appPart = package.ExtendedFilePropertiesPart ?? package.AddExtendedFilePropertiesPart();
            appPart.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
            appPart.Properties.DigitalSignature = new DigitalSignature();
            appPart.Properties.Save();
        }

        private static X509Certificate2 CreateSelfSignedSigningCertificate() {
            using RSA rsa = RSA.Create(2048);
            var request = new CertificateRequest(
                "CN=OfficeIMO Package Signing Test",
                rsa,
                HashAlgorithmName.SHA256,
                RSASignaturePadding.Pkcs1);

            request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: false));

            using X509Certificate2 certificate = request.CreateSelfSigned(
                DateTimeOffset.UtcNow.AddDays(-1),
                DateTimeOffset.UtcNow.AddDays(1));

            return new X509Certificate2(certificate.Export(X509ContentType.Pfx), (string?)null, X509KeyStorageFlags.Exportable);
        }

        private static void AddCertificateToCurrentUserStore(X509Certificate2 certificate) {
            using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadWrite);
            store.Add(certificate);
        }

        private static void RemoveCertificateFromCurrentUserStore(string? thumbprint) {
            if (string.IsNullOrWhiteSpace(thumbprint)) {
                return;
            }

            using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadWrite);
            foreach (X509Certificate2 certificate in store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, validOnly: false)) {
                try {
                    store.Remove(certificate);
                } finally {
                    certificate.Dispose();
                }
            }
        }
    }
}

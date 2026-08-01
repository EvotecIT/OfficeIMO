using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Collections.Generic;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_DigitalSignature_PackageAndMacroProjectCapabilitiesAreExplicitlySeparate() {
            Assert.True(WordSigningCapabilities.Package.IsSupported);
            Assert.Equal(WordSigningCapabilityKind.OpcPackage, WordSigningCapabilities.Package.Kind);
            Assert.Equal(
                System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows),
                WordSigningCapabilities.MacroProject.IsSupported);
            Assert.Equal(WordSigningCapabilityKind.VbaMacroProject, WordSigningCapabilities.MacroProject.Kind);
        }

        [Fact]
        public void Test_DigitalSignature_RequiredIndeterminateRevocationIsNotValidUnderPolicy() {
            var signaturePart = new WordSignaturePartInfo(
                "/_xmlsignatures/sig1.xml",
                "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
                null,
                0,
                null,
                Array.Empty<string>(),
                Array.Empty<WordSignatureReferenceInfo>(),
                Array.Empty<WordSignatureTimestampInfo>(),
                Array.Empty<string>(),
                null,
                Array.Empty<string>());
            var result = new WordSignaturePartValidationResult(
                signaturePart,
                WordSignatureValidationState.Passed,
                WordSignatureValidationState.Passed,
                WordSignatureValidationState.NotChecked,
                revocationCheckRequired: true,
                WordSignatureValidationState.NotPresent,
                null,
                Array.Empty<OfficeIMO.Security.Rfc3161TimestampVerificationResult>(),
                Array.Empty<WordSignatureValidationFinding>());

            Assert.True(result.RevocationCheckRequired);
            Assert.False(result.IsValidUnderPolicy);
        }

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
        public void Test_DigitalSignature_UnsignedValidationReturnsBeforeApplyingSignedPackageByteBudget() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph(new string('x', 4096));

            WordSignatureValidationReport validation = document.ValidateSignatures(new WordSignatureValidationOptions {
                MaxPackageBytes = 1
            });

            Assert.False(validation.HasSignatures);
            Assert.Equal(WordSignatureValidationState.NotPresent, validation.PackageStructureStatus);
            Assert.DoesNotContain(validation.Diagnostics, finding => finding.Code == "PackageByteLimitExceeded");
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
                document.Save();
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
                Assert.Equal(
                    "/word/document.xml?ContentType=application%2Fvnd.openxmlformats-officedocument.wordprocessingml.document.main%2Bxml",
                    signedReference.Uri);
                Assert.Equal("http://www.w3.org/2001/04/xmlenc#sha256", signedReference.DigestMethodAlgorithm);
                Assert.True(signedReference.HasDigestValue);
                Assert.Equal("T2ZmaWNlSU1P", signedReference.DigestValue);
                Assert.True(signedReference.IsPackagePartReference);
                Assert.Equal("/word/document.xml", signedReference.TargetPartUri);
                Assert.True(signedReference.TargetPartExists);
                Assert.Contains("CN=OfficeIMO Test", signaturePart.X509SubjectNames);
                Assert.Empty(signatures.UnsupportedDetails);

                WordSignatureValidationReport validation = document.ValidateSignatures();
                Assert.False(validation.IsStructurallyValid);
                Assert.Equal(WordSignatureValidationState.Passed, validation.PackageStructureStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.XmlSignatureStatus);
                Assert.Equal(WordSignatureValidationState.Unsupported, validation.CryptographicStatus);
                Assert.Equal(WordSignatureValidationState.NotPresent, validation.CertificateChainStatus);
                Assert.Equal(WordSignatureValidationState.NotPresent, validation.RevocationStatus);
                Assert.Equal(WordSignatureValidationState.NotPresent, validation.TimestampStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.SignedPartCoverageStatus);
                Assert.Equal(WordSignatureValidationState.Failed, validation.SignedPartDigestStatus);
                Assert.Contains(validation.Findings, finding => finding.Contains("digest did not match", System.StringComparison.OrdinalIgnoreCase));
                Assert.Contains(validation.Diagnostics, finding => finding.Code == "SignerCertificateMissing");
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
                document.Save();
            }

            AddDigitalSignatureMetadata(
                filePath,
                CreateSignatureXml(
                    includeOpcSignatureTime: true,
                    opcSignatureTimeValue: opcTimestampValue,
                    includeXadesSigningTime: true,
                    xadesSigningTimeValue: xadesTimestampValue));

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
                Assert.Equal(WordSignatureValidationState.NotPresent, validation.TimestampStatus);
                Assert.Contains(validation.Diagnostics, finding => finding.Code == "ClaimedSigningTimeNotTrusted");
                Assert.Contains(validation.SignatureInfo.Details, detail => detail.Contains("Signature timestamp", System.StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesSupportsSignedFixture() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "DigitalSignatures", "signed-valid.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing signed DOCX fixture: {sourcePath}");

            using (WordDocument document = WordDocument.Load(sourcePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                WordSignatureInfo signatures = document.InspectSignatures();

                Assert.True(signatures.HasSignatures);
                Assert.True(signatures.HasDigitalSignatureOriginPart);
                Assert.NotEmpty(signatures.SignatureParts);
                Assert.Contains(signatures.SignatureParts.SelectMany(part => part.X509SubjectNames), subject =>
                    subject.Contains("OfficeIMO Fixture Package Signing", System.StringComparison.OrdinalIgnoreCase));

                var validationOptions = new WordSignatureValidationOptions();
                validationOptions.CertificateValidation.DisableCertificateDownloads = false;
                validationOptions.CertificateValidation.ChainEvaluator = static (_, _) => true;
                WordSignatureValidationReport validation = document.ValidateSignatures(validationOptions);

                Assert.True(validation.IsStructurallyValid, string.Join(System.Environment.NewLine, validation.Findings));
                Assert.Equal(WordSignatureValidationState.Passed, validation.PackageStructureStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.XmlSignatureStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.SignedPartCoverageStatus);
                Assert.NotEqual(WordSignatureValidationState.Failed, validation.SignedPartDigestStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.CryptographicStatus);
                Assert.Equal(WordSignatureValidationState.Failed, validation.CertificateChainStatus);
                Assert.Equal(WordSignatureValidationState.NotChecked, validation.RevocationStatus);
                Assert.Equal(WordSignatureValidationState.NotPresent, validation.TimestampStatus);
                Assert.False(validation.IsValidUnderPolicy);
                Assert.Contains(validation.Diagnostics, finding => finding.Code == "CertificateEnhancedKeyUsageInvalid");
                Assert.Contains(validation.SignatureInfo.SignatureParts.SelectMany(part => part.Timestamps), timestamp =>
                    !string.IsNullOrWhiteSpace(timestamp.Value));
                Assert.Contains(validation.SignatureInfo.SignatureParts.SelectMany(part => part.SignedReferences), reference =>
                    reference.HasDigestValue &&
                    reference.IsPackagePartReference &&
                    reference.TargetPartExists == true);
            }
        }

        [Fact]
        public void Test_DigitalSignature_DefaultCertificatePolicyDisablesIssuerDownloadsAndAllowsOptIn() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "DigitalSignatures", "signed-valid.docx"));
            using WordDocument document = WordDocument.Load(sourcePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });

            bool defaultChainEvaluated = false;
            bool? defaultDisableCertificateDownloads = null;
            var defaultOptions = new WordSignatureValidationOptions();
            defaultOptions.CertificateValidation.ChainEvaluator = (_, chain) => {
                defaultChainEvaluated = true;
                defaultDisableCertificateDownloads = ReadDisableCertificateDownloads(chain.ChainPolicy);
                return true;
            };
            WordSignatureValidationReport defaultValidation = document.ValidateSignatures(defaultOptions);

            bool optInChainEvaluated = false;
            bool? optedInDisableCertificateDownloads = null;
            var optInOptions = new WordSignatureValidationOptions();
            optInOptions.CertificateValidation.DisableCertificateDownloads = false;
            optInOptions.CertificateValidation.ChainEvaluator = (_, chain) => {
                optInChainEvaluated = true;
                optedInDisableCertificateDownloads = ReadDisableCertificateDownloads(chain.ChainPolicy);
                return true;
            };
            document.ValidateSignatures(optInOptions);

            if (typeof(X509ChainPolicy).GetProperty("DisableCertificateDownloads") != null) {
                Assert.True(defaultChainEvaluated);
                Assert.True(optInChainEvaluated);
                Assert.True(defaultDisableCertificateDownloads);
                Assert.False(optedInDisableCertificateDownloads);
            } else {
                Assert.False(defaultChainEvaluated);
                Assert.True(optInChainEvaluated);
                Assert.Contains(defaultValidation.Diagnostics, finding =>
                    finding.Code == "CertificateDownloadPolicyUnavailable");
            }
        }

        [Fact]
        public void Test_DigitalSignature_MetadataInspectionDefersTransformAwareDigestsToValidation() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "DigitalSignatures", "signed-valid.docx"));
            using WordDocument document = WordDocument.Load(sourcePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });

            WordSignatureInfo inspection = document.InspectSignatures();
            WordSignatureReferenceInfo[] inspectedTransformedReferences = inspection.SignatureParts
                .SelectMany(part => part.SignedReferences)
                .Where(reference => reference.TransformAlgorithms.Count > 0)
                .ToArray();

            Assert.NotEmpty(inspectedTransformedReferences);
            Assert.All(inspectedTransformedReferences, reference =>
                Assert.Equal(WordSignatureValidationState.NotChecked, reference.DigestVerificationStatus));

            WordSignatureValidationReport validation = document.ValidateSignatures();
            WordSignatureReferenceInfo[] validatedTransformedReferences = validation.SignatureInfo.SignatureParts
                .SelectMany(part => part.SignedReferences)
                .Where(reference => reference.TransformAlgorithms.Count > 0)
                .ToArray();
            Assert.NotEmpty(validatedTransformedReferences);
            Assert.All(validatedTransformedReferences, reference =>
                Assert.Equal(WordSignatureValidationState.Passed, reference.DigestVerificationStatus));
        }

        [Fact]
        public void Test_DigitalSignature_SharedPackageInspectorReadsOpenXmlSignatureMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureSharedInspector.docx");
            byte[] signatureBytes = CreateSignatureXml();

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Shared package signature inspector carrier");
                document.Save();
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, false)) {
                OfficePackageSignatureInfo signatures = OfficePackageSignatureInspector.Inspect(
                    package,
                    package.DigitalSignatureOriginPart,
                    package.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null,
                    File.ReadAllBytes(filePath));

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
                Assert.Empty(signatures.UnsupportedDetails);
            }
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesVerifiesSimplePackagePartDigest() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureValidSimpleDigest.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Simple digest verification carrier");
                document.Save();
            }

            AddDigitalSignatureMetadata(filePath, CreateSignatureXml(digestValue: ComputePackagePartSha256Digest(filePath, "/word/document.xml")));

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
                document.Save();
            }

            AddDigitalSignatureMetadata(filePath, CreateSignatureXml(digestValue: "T2ZmaWNlSU1P"));

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                WordSignatureValidationReport validation = document.ValidateSignatures();

                WordSignatureReferenceInfo signedReference = Assert.Single(Assert.Single(validation.SignatureInfo.SignatureParts).SignedReferences);
                Assert.Equal(WordSignatureValidationState.Failed, signedReference.DigestVerificationStatus);
                Assert.Equal(WordSignatureValidationState.Failed, validation.SignedPartDigestStatus);
                Assert.Contains(validation.Findings, finding => finding.Contains("digest did not match", System.StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_DigitalSignature_DigestValidationRejectsSignedContentTypeMismatch() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureContentTypeBinding.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Content type binding");
                document.Save();
            }

            XNamespace ds = "http://www.w3.org/2000/09/xmldsig#";
            using var archive = new OfficePackageSignatureArchive(File.ReadAllBytes(filePath));
            var reference = new XElement(ds + "Reference",
                new XAttribute("URI", "/word/document.xml?ContentType=application%2Fvnd.openxmlformats-officedocument.wordprocessingml.document.main%2Bxml"),
                new XElement(ds + "DigestMethod", new XAttribute("Algorithm", "http://www.w3.org/2001/04/xmlenc#sha256")));
            reference.Add(new XElement(ds + "DigestValue", archive.ComputeDigestValue(reference, 16 * 1024 * 1024)));
            reference.SetAttributeValue("URI", "/word/document.xml?ContentType=application%2Foctet-stream");

            OfficePackageDigestResult result = archive.VerifyReference(reference, 16 * 1024 * 1024);

            Assert.Equal(OfficePackageSignatureDigestVerificationStatus.Failed, result.Status);
            Assert.Contains("content type", result.Detail, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_DigitalSignature_FragmentBearingPackageReferenceIsUnsupported() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureFragmentReference.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Fragment-bearing package reference");
                document.Save();
            }
            string digest = ComputePackagePartSha256Digest(filePath, "/word/document.xml");
            AddDigitalSignatureMetadata(filePath, CreateSignatureXml(
                referenceUri: "/word/document.xml#target",
                digestValue: digest));

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures();

            Assert.Equal(WordSignatureValidationState.Unsupported, validation.SignedPartCoverageStatus);
            Assert.NotEqual(WordSignatureValidationState.Passed, validation.SignedPartDigestStatus);
            Assert.False(validation.IsValidUnderPolicy);
        }

        [Fact]
        public void Test_DigitalSignature_UnsupportedManifestReferenceCannotPassCoveragePolicy() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureUnsupportedManifestReference.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Unsupported manifest reference");
                document.Save();
            }
            string documentDigest = ComputePackagePartSha256Digest(filePath, "/word/document.xml");
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\">" +
                "<SignedInfo><SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                "<Reference URI=\"#manifest\"><DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>T2ZmaWNlSU1P</DigestValue></Reference></SignedInfo>" +
                "<Object><Manifest Id=\"manifest\">" +
                "<Reference URI=\"/word/document.xml\"><DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>" + documentDigest + "</DigestValue></Reference>" +
                "<Reference URI=\"https://example.invalid/external\"><DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>T2ZmaWNlSU1P</DigestValue></Reference>" +
                "</Manifest></Object></Signature>");
            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures();

            Assert.Equal(WordSignatureValidationState.Unsupported, validation.SignedPartCoverageStatus);
            Assert.Equal(WordSignatureValidationState.Unsupported, validation.SignedPartDigestStatus);
            Assert.False(validation.IsValidUnderPolicy);
        }

        [Fact]
        public void Test_DigitalSignature_IgnoresPackageReferencesInUnsignedManifestObjects() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureUnsignedManifest.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Unsigned manifest must not establish package coverage");
                document.Save();
            }
            string documentDigest = ComputePackagePartSha256Digest(filePath, "/word/document.xml");
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\">" +
                "<SignedInfo><SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                "<Reference URI=\"#signed-object\"><DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>T2ZmaWNlSU1P</DigestValue></Reference></SignedInfo>" +
                "<Object Id=\"signed-object\"><SignatureProperties /></Object>" +
                "<Object><Manifest><Reference URI=\"/word/document.xml\"><DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>" + documentDigest + "</DigestValue></Reference></Manifest></Object>" +
                "</Signature>");
            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures();

            WordSignaturePartInfo part = Assert.Single(validation.SignatureInfo.SignatureParts);
            Assert.DoesNotContain(part.SignedReferences, reference => reference.IsPackagePartReference);
            Assert.Equal(WordSignatureValidationState.Unsupported, validation.SignedPartCoverageStatus);
            Assert.Contains(part.UnsupportedDetails, detail => detail.Contains("not authenticated by SignedInfo", System.StringComparison.Ordinal));
            Assert.False(validation.IsValidUnderPolicy);
        }

        [Fact]
        public void Test_DigitalSignature_EmptyUriReferenceDoesNotAuthenticateExcludedManifest() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureEmptyUriManifest.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Empty URI excludes the signature subtree");
                document.Save();
            }
            string documentDigest = ComputePackagePartSha256Digest(filePath, "/word/document.xml");
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\">" +
                "<SignedInfo><SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                "<Reference URI=\"\"><Transforms><Transform Algorithm=\"http://www.w3.org/2000/09/xmldsig#enveloped-signature\" /></Transforms>" +
                "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>T2ZmaWNlSU1P</DigestValue></Reference></SignedInfo>" +
                "<Object><Manifest><Reference URI=\"/word/document.xml\"><DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>" + documentDigest + "</DigestValue></Reference></Manifest></Object>" +
                "</Signature>");
            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures();

            WordSignaturePartInfo part = Assert.Single(validation.SignatureInfo.SignatureParts);
            Assert.DoesNotContain(part.SignedReferences, reference => reference.IsPackagePartReference);
            Assert.Equal(WordSignatureValidationState.Unsupported, validation.SignedPartCoverageStatus);
            Assert.Contains(part.UnsupportedDetails, detail => detail.Contains("not authenticated by SignedInfo", System.StringComparison.Ordinal));
            Assert.False(validation.IsValidUnderPolicy);
        }

        [Fact]
        public void Test_DigitalSignature_FragmentTransformCannotAuthenticateExcludedNestedManifest() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureTransformedFragmentManifest.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Fragment transforms must preserve authenticated manifests");
                document.Save();
            }
            string documentDigest = ComputePackagePartSha256Digest(filePath, "/word/document.xml");
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\">" +
                "<SignedInfo><SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                "<Reference URI=\"#signed-object\"><Transforms><Transform Algorithm=\"http://www.w3.org/TR/1999/REC-xpath-19991116\">" +
                "<XPath>not(ancestor-or-self::Manifest)</XPath></Transform></Transforms>" +
                "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>T2ZmaWNlSU1P</DigestValue></Reference></SignedInfo>" +
                "<Object Id=\"signed-object\"><Manifest><Reference URI=\"/word/document.xml\">" +
                "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>" + documentDigest + "</DigestValue>" +
                "</Reference></Manifest><SignatureProperties /></Object></Signature>");
            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures();

            WordSignaturePartInfo part = Assert.Single(validation.SignatureInfo.SignatureParts);
            Assert.DoesNotContain(part.SignedReferences, reference => reference.IsPackagePartReference);
            Assert.Contains(part.UnsupportedDetails, detail => detail.Contains("transform", System.StringComparison.OrdinalIgnoreCase));
            Assert.NotEqual(WordSignatureValidationState.Passed, validation.SignedPartCoverageStatus);
            Assert.False(validation.IsValidUnderPolicy);
        }

        [Fact]
        public void Test_DigitalSignature_BoundsAggregateDigestWorkBeforeReadingDuplicateParts() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureDigestWorkBudget.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph(new string('x', 4096));
                document.Save();
            }
            string documentDigest = ComputePackagePartSha256Digest(filePath, "/word/document.xml");
            string reference = "<Reference URI=\"/word/document.xml\"><DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>" + documentDigest + "</DigestValue></Reference>";
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo>" +
                "<SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                reference + reference +
                "</SignedInfo></Signature>");
            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(new WordSignatureValidationOptions {
                MaxTotalDigestBytes = 1
            });

            WordSignaturePartInfo part = Assert.Single(validation.SignatureInfo.SignatureParts);
            Assert.True(part.HasParseError);
            Assert.Empty(part.SignedReferences);
            Assert.Contains(part.UnsupportedDetails, detail => detail.Contains("aggregate digest-work limit", System.StringComparison.Ordinal));
            Assert.False(validation.IsValidUnderPolicy);
        }

        [Fact]
        public void Test_DigitalSignature_ChargesEveryPackageTransformPassToDigestWorkBudget() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureTransformDigestWorkBudget.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph(new string('x', 4096));
                document.Save();
            }
            long documentPartLength;
            using (ZipArchive archive = ZipFile.OpenRead(filePath)) {
                documentPartLength = archive.GetEntry("word/document.xml")!.Length;
            }
            string transforms =
                "<Transforms>" +
                "<Transform Algorithm=\"" + SignedXml.XmlDsigC14NTransformUrl + "\" />" +
                "<Transform Algorithm=\"" + SignedXml.XmlDsigC14NTransformUrl + "\" />" +
                "</Transforms>";
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo>" +
                "<SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                "<Reference URI=\"/word/document.xml\">" + transforms +
                "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>T2ZmaWNlSU1P</DigestValue></Reference>" +
                "</SignedInfo></Signature>");
            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(new WordSignatureValidationOptions {
                MaxTotalDigestBytes = checked(documentPartLength * 2)
            });

            WordSignaturePartInfo part = Assert.Single(validation.SignatureInfo.SignatureParts);
            Assert.True(part.HasParseError);
            Assert.Empty(part.SignedReferences);
            Assert.Contains(part.UnsupportedDetails, detail => detail.Contains("aggregate digest-work limit", StringComparison.Ordinal));
            Assert.False(validation.IsValidUnderPolicy);
        }

        [Fact]
        public void Test_DigitalSignature_CountsAllSignedInfoReferencesBeforeCryptographicValidation() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureSignedInfoReferenceLimit.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("SignedInfo reference count resource limit");
                document.Save();
            }
            string documentDigest = ComputePackagePartSha256Digest(filePath, "/word/document.xml");
            string digest = "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>T2ZmaWNlSU1P</DigestValue>";
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo>" +
                "<SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                "<Reference URI=\"/word/document.xml\"><DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>" + documentDigest + "</DigestValue></Reference>" +
                "<Reference URI=\"#payload\">" + digest + "</Reference>" +
                "</SignedInfo><Object Id=\"payload\">payload</Object></Signature>");
            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(new WordSignatureValidationOptions {
                MaxSignedReferences = 1,
                ValidateCryptographicSignature = false
            });

            WordSignaturePartInfo part = Assert.Single(validation.SignatureInfo.SignatureParts);
            Assert.True(part.HasParseError);
            Assert.Contains(part.UnsupportedDetails, detail => detail.Contains("more than 1 authenticated references", System.StringComparison.Ordinal));
            Assert.Contains(validation.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
        }

        [Fact]
        public void Test_DigitalSignature_ArchiveDoesNotTreatZipDirectoryEntriesAsPackageParts() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureDirectoryEntry.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Explicit ZIP directory entry");
                document.Save();
            }
            using (ZipArchive archive = ZipFile.Open(filePath, ZipArchiveMode.Update)) {
                archive.CreateEntry("custom/");
            }

            using var package = new OfficePackageSignatureArchive(File.ReadAllBytes(filePath));

            Assert.False(package.ContainsPart("/custom/"));
            Assert.DoesNotContain(package.PartUris, uri => uri.EndsWith("/", System.StringComparison.Ordinal));
            Assert.DoesNotContain("/custom/", package.PartUris);
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesLeavesTransformedDigestVerificationUnsupported() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureTransformedDigestUnsupported.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Transformed digest verification carrier");
                document.Save();
            }

            AddDigitalSignatureMetadata(
                filePath,
                CreateSignatureXml(
                    digestValue: ComputePackagePartSha256Digest(filePath, "/word/document.xml"),
                    transformAlgorithm: "http://www.w3.org/2000/09/xmldsig#enveloped-signature"));

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                WordSignatureValidationReport validation = document.ValidateSignatures();

                WordSignatureReferenceInfo signedReference = Assert.Single(Assert.Single(validation.SignatureInfo.SignatureParts).SignedReferences);
                Assert.Equal(WordSignatureValidationState.Unsupported, signedReference.DigestVerificationStatus);
                Assert.Equal("http://www.w3.org/2000/09/xmldsig#enveloped-signature", Assert.Single(signedReference.TransformAlgorithms));
                Assert.Equal(WordSignatureValidationState.Unsupported, validation.SignedPartDigestStatus);
                Assert.Contains(validation.Findings, finding => finding.Contains("transform", System.StringComparison.OrdinalIgnoreCase));
            }
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesReportsMissingReferenceDigestValue() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureMissingDigestValue.docx");
            byte[] signatureBytes = CreateSignatureXml(includeDigestValue: false);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Missing reference digest value carrier");
                document.Save();
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
                document.Save();
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
                document.Save();
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                WordSignatureValidationReport validation = document.ValidateSignatures();

                Assert.False(validation.IsStructurallyValid);
                Assert.Equal(WordSignatureValidationState.Passed, validation.PackageStructureStatus);
                Assert.Equal(WordSignatureValidationState.Failed, validation.XmlSignatureStatus);
                Assert.Contains(validation.Findings, finding => finding.Contains("could not be parsed", System.StringComparison.OrdinalIgnoreCase));
                Assert.True(Assert.Single(validation.SignatureInfo.SignatureParts).HasParseError);
            }
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesReportsOversizedSignatureWithoutThrowing() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureOversized.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Oversized signature resource limit");
                document.Save();
            }
            AddDigitalSignatureMetadata(filePath, Encoding.UTF8.GetBytes("<Signature>" + new string('x', 2048) + "</Signature>"));

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(new WordSignatureValidationOptions { MaxSignatureBytes = 256 });

            Assert.False(validation.IsValidUnderPolicy);
            Assert.Contains(validation.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
        }

        [Fact]
        public void Test_DigitalSignature_ValidationRejectsSignatureCountBeforeParsingParts() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureCount.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signature count resource limit");
                document.Save();
            }
            AddDigitalSignatureMetadata(filePath, Encoding.UTF8.GetBytes("not xml"), signatureCount: 2);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(new WordSignatureValidationOptions { MaxSignatureParts = 1 });

            Assert.False(validation.IsValidUnderPolicy);
            Assert.Empty(validation.SignatureInfo.SignatureParts);
            Assert.Contains(validation.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
            Assert.DoesNotContain(validation.Findings, finding => finding.Contains("could not be parsed", System.StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Test_DigitalSignature_ArchivePartLimitStopsInspectionBeforePackageTraversal() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureArchivePartLimit.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Bounded archive failure must stop digest work");
                document.Save();
            }
            AddDigitalSignatureMetadata(filePath, CreateSignatureXml(
                digestValue: ComputePackagePartSha256Digest(filePath, "/word/document.xml")));

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(new WordSignatureValidationOptions {
                MaxPackageParts = 1
            });

            Assert.Empty(validation.SignatureInfo.SignatureParts);
            Assert.Contains(validation.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");

            WordSignatureValidationReport permissive = loaded.ValidateSignatures(new WordSignatureValidationOptions {
                MaxPackageParts = 10000
            });
            Assert.Single(permissive.SignatureInfo.SignatureParts);
        }

        [Fact]
        public void Test_DigitalSignature_SigningReadbackReusesConfiguredResourceLimits() {
            var signingOptions = new WordPackageSigningOptions {
                MaxPackageBytes = 768L * 1024 * 1024,
                MaxPackageParts = 12000,
                MaxPartBytes = 384L * 1024 * 1024,
                MaxTotalDigestBytes = 640L * 1024 * 1024,
                MaxSignedReferences = 8000,
                MaxSignatureBytes = 24L * 1024 * 1024,
                MaxCertificates = 96,
                MaxCertificateBytes = 6L * 1024 * 1024,
                MaxTotalCertificateBytes = 80L * 1024 * 1024
            };

            WordSignatureValidationOptions validationOptions = WordDocument.CreateSigningReadbackOptions(
                signingOptions,
                signatureCount: 48);

            Assert.Equal(signingOptions.MaxPackageBytes, validationOptions.MaxPackageBytes);
            Assert.Equal(signingOptions.MaxPackageParts, validationOptions.MaxPackageParts);
            Assert.Equal(signingOptions.MaxPartBytes, validationOptions.MaxPartBytes);
            Assert.Equal(signingOptions.MaxTotalDigestBytes, validationOptions.MaxTotalDigestBytes);
            Assert.Equal(signingOptions.MaxSignedReferences, validationOptions.MaxSignedReferences);
            Assert.Equal(signingOptions.MaxSignatureBytes, validationOptions.MaxSignatureBytes);
            Assert.Equal(signingOptions.MaxCertificates, validationOptions.MaxCertificates);
            Assert.Equal(signingOptions.MaxCertificateBytes, validationOptions.MaxCertificateBytes);
            Assert.Equal(signingOptions.MaxTotalCertificateBytes * 48, validationOptions.MaxTotalCertificateBytes);
            Assert.Equal(48, validationOptions.MaxSignatureParts);
        }

        [Fact]
        public void Test_DigitalSignature_SigningRejectsReferenceCountOutsideConfiguredLimitAtomically() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureReferenceLimit.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signing reference limit");
                document.Save();
            }
            byte[] originalBytes = File.ReadAllBytes(filePath);

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult result = WordDocument.TrySignPackage(
                filePath,
                certificate,
                new WordPackageSigningOptions { MaxSignedReferences = 1 });

            Assert.False(result.Succeeded);
            Assert.Contains(result.Details, detail => detail.Contains("authenticated references", System.StringComparison.OrdinalIgnoreCase));
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
        }

        [Fact]
        public void Test_DigitalSignature_SigningChargesRelationshipReferencesAgainstConfiguredLimit() {
            string sourcePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureRelationshipReferenceSource.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Relationship reference budget");
                document.Save();
            }
            string probePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureRelationshipReferenceProbe.docx");
            string boundaryPath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureRelationshipReferenceBoundary.docx");
            string exactPath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureRelationshipReferenceExact.docx");
            File.Copy(sourcePath, probePath);
            File.Copy(sourcePath, boundaryPath);
            File.Copy(sourcePath, exactPath);

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            var probeOptions = new WordPackageSigningOptions {
                PartUris = new[] { "/word/document.xml" },
                IncludePackageRelationships = true,
                IncludePartRelationships = false
            };
            WordPackageSigningResult probe = WordDocument.SignPackage(probePath, certificate, probeOptions);
            Assert.True(probe.Succeeded);
            Assert.True(probe.SignedRelationshipSelectorCount > 0);
            int limitWithoutRelationshipReference = checked(
                probe.SignedPartCount + probe.SignedRelationshipSelectorCount + 1);

            WordPackageSigningResult boundary = WordDocument.TrySignPackage(
                boundaryPath,
                certificate,
                new WordPackageSigningOptions {
                    PartUris = probeOptions.PartUris,
                    IncludePackageRelationships = true,
                    IncludePartRelationships = false,
                    MaxSignedReferences = limitWithoutRelationshipReference
                });
            WordPackageSigningResult exact = WordDocument.TrySignPackage(
                exactPath,
                certificate,
                new WordPackageSigningOptions {
                    PartUris = probeOptions.PartUris,
                    IncludePackageRelationships = true,
                    IncludePartRelationships = false,
                    MaxSignedReferences = checked(limitWithoutRelationshipReference + 1)
                });

            Assert.False(boundary.Succeeded);
            Assert.Contains(boundary.Details, detail => detail.Contains("authenticated references", StringComparison.OrdinalIgnoreCase));
            Assert.True(exact.Succeeded, string.Join(Environment.NewLine, exact.Details));
        }

        [Theory]
        [InlineData("idPackageObject")]
        [InlineData("idSignatureTime")]
        public void Test_DigitalSignature_SigningRejectsReservedInternalSignatureIdsAtomically(string signatureId) {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureReservedId.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Reserved signature identifiers");
                document.Save();
            }
            byte[] originalBytes = File.ReadAllBytes(filePath);

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult result = WordDocument.TrySignPackage(
                filePath,
                certificate,
                new WordPackageSigningOptions { SignatureId = signatureId });

            Assert.False(result.Succeeded);
            Assert.Contains(result.Details, detail => detail.Contains("reserved", StringComparison.OrdinalIgnoreCase));
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
            using WordprocessingDocument preserved = WordprocessingDocument.Open(filePath, false);
            Assert.Null(preserved.DigitalSignatureOriginPart);
        }

        [Fact]
        public void Test_DigitalSignature_SigningRejectsCertificateCountOutsideConfiguredLimitAtomically() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureSigningCertificateLimit.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signing certificate limit");
                document.Save();
            }
            byte[] originalBytes = File.ReadAllBytes(filePath);

            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            using X509Certificate2 additional = CreateSelfSignedSigningCertificate("CN=OfficeIMO Additional Certificate");
            WordPackageSigningResult result = WordDocument.TrySignPackage(
                filePath,
                signer,
                new WordPackageSigningOptions {
                    AdditionalCertificates = new[] { additional },
                    MaxCertificates = 1
                });

            Assert.False(result.Succeeded);
            Assert.Contains(result.Details, detail => detail.Contains("certificate limit", System.StringComparison.OrdinalIgnoreCase));
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
        }

        [Fact]
        public void Test_DigitalSignature_SigningRejectsFinalPackageOutsideConfiguredByteLimitAtomically() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureFinalPackageByteLimit.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Final signed-package byte limit");
                document.Save();
            }
            byte[] originalBytes = File.ReadAllBytes(filePath);

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult result = WordDocument.TrySignPackage(
                filePath,
                certificate,
                new WordPackageSigningOptions { MaxPackageBytes = originalBytes.LongLength });

            Assert.False(result.Succeeded);
            Assert.Contains(result.Details, detail => detail.Contains("signed package exceeds", System.StringComparison.OrdinalIgnoreCase));
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
        }

        [Fact]
        public void Test_DigitalSignature_ValidationCountsDeclaredEmbeddedCertificates() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureCertificateCount.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Certificate count resource limit");
                document.Save();
            }
            AddDigitalSignatureMetadata(filePath, CreateSignatureXmlWithCertificates("!", "!"));

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(new WordSignatureValidationOptions { MaxCertificates = 1 });

            Assert.Contains(validation.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
            Assert.Contains("certificate limit", Assert.Single(validation.SignatureInfo.SignatureParts).ParseError!, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_DigitalSignature_ValidationBoundsEmbeddedCertificateBeforeDecoding() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureCertificateBytes.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Certificate byte resource limit");
                document.Save();
            }
            AddDigitalSignatureMetadata(filePath, CreateSignatureXmlWithCertificates(new string('A', 256)));

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(new WordSignatureValidationOptions { MaxCertificateBytes = 16 });

            Assert.Contains(validation.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
            Assert.Contains("byte limit", Assert.Single(validation.SignatureInfo.SignatureParts).ParseError!, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_DigitalSignature_ValidationBoundsAggregateCertificateWorkAcrossSignatureParts() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureAggregateCertificateBytes.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Aggregate certificate byte resource limit");
                document.Save();
            }
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            byte[] certificateBytes = certificate.Export(X509ContentType.Cert);
            AddDigitalSignatureMetadata(
                filePath,
                CreateSignatureXmlWithCertificates(System.Convert.ToBase64String(certificateBytes)),
                signatureCount: 2);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(new WordSignatureValidationOptions {
                MaxTotalCertificateBytes = certificateBytes.LongLength + 1
            });

            Assert.Contains(validation.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
            Assert.Contains(validation.SignatureInfo.SignatureParts, part =>
                part.ParseError?.Contains("aggregate certificate limit", System.StringComparison.OrdinalIgnoreCase) == true);
        }

        [Fact]
        public void Test_DigitalSignature_ValidationIgnoresCertificateNamedBusinessDataOutsideKeyInfo() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureCertificateBusinessData.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Certificate-like business data");
                document.Save();
            }
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            string signatureXml = Encoding.UTF8.GetString(CreateSignatureXmlWithCertificates(
                System.Convert.ToBase64String(certificate.Export(X509ContentType.Cert))));
            signatureXml = signatureXml.Replace(
                "<SignatureValue>",
                "<Object><X509Certificate>!</X509Certificate><X509Certificate>!</X509Certificate></Object><SignatureValue>");
            AddDigitalSignatureMetadata(filePath, Encoding.UTF8.GetBytes(signatureXml));

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(new WordSignatureValidationOptions { MaxCertificates = 1 });

            Assert.Null(Assert.Single(validation.SignatureInfo.SignatureParts).ParseError);
            Assert.DoesNotContain(validation.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesReportsApplicationMetadataWithoutOriginAsUnsupported() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureApplicationOnly.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Application signature metadata only");
                document.ApplicationProperties.DigitalSignature = new DigitalSignature();
                document.Save(new WordSaveOptions { SignedDocumentPolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation });
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
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
                document.Save();
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.Save(new WordSaveOptions { SignedDocumentPolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation });
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

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                WordFeatureFinding signatures = Assert.Single(document.InspectFeatures().FindFeatures("Digital signatures"));

                Assert.Equal(WordFeatureSupportLevel.PartiallyEditable, signatures.SupportLevel);
                Assert.Contains(signatures.Details, detail => detail.Contains("origin.sigs", System.StringComparison.OrdinalIgnoreCase));
                Assert.Contains(signatures.Details, detail => detail.Contains("_xmlsignatures", System.StringComparison.OrdinalIgnoreCase));
                Assert.Contains("validated", signatures.Note, System.StringComparison.OrdinalIgnoreCase);
                Assert.Contains("macro-project signing", signatures.Note, System.StringComparison.OrdinalIgnoreCase);
            }
        }

        [Fact]
        public void Test_DigitalSignature_SaveBlocksSignedDocumentByDefault() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureSaveBlocked.docx");
            byte[] signatureBytes = CreateSignatureXml();

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed blocked save");
                document.Save();
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AddParagraph("Mutation after signing");

                WordSignatureSavePolicyException exception = Assert.Throws<WordSignatureSavePolicyException>(() => document.Save());

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
                document.Save();
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AddParagraph("Mutation after signing");
                document.Save(new WordSaveOptions { SignedDocumentPolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation });
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Contains(document.Paragraphs, paragraph => paragraph.Text == "Mutation after signing");
                Assert.True(document.InspectSignatures().HasSignatures);
            }
        }

        [Fact]
        public void Test_DigitalSignature_ToDocxStreamBlocksSignedDocumentByDefault() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureStreamBlocked.docx");
            byte[] signatureBytes = CreateSignatureXml();

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed stream blocked");
                document.Save();
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Throws<WordSignatureSavePolicyException>(() => document.ToStream());
                using MemoryStream stream = document.ToStream(options:
                    new WordSaveOptions { SignedDocumentPolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation });

                Assert.True(stream.Length > 0);
            }
        }

        [Fact]
        public void Test_DigitalSignature_DisposeSignedAutoSavePropagatesPolicyFailureWithoutSaving() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureAutoSaveDispose.docx");
            byte[] signatureBytes = CreateSignatureXml();

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed autosave source");
                document.Save();
            }

            AddDigitalSignatureMetadata(filePath, signatureBytes);

            WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose
            });
            loaded.AddParagraph("Mutation that should not autosave");
            Assert.Throws<WordSignatureSavePolicyException>(() => loaded.Dispose());

            Assert.False(filePath.IsFileLocked());
            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.DoesNotContain(document.Paragraphs, paragraph => paragraph.Text == "Mutation that should not autosave");
                Assert.True(document.InspectSignatures().HasSignatures);
            }
        }

        [Fact]
        public void Test_DigitalSignature_TrySignPackageRejectsInvalidThumbprintCharacters() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureInvalidThumbprint.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Invalid thumbprint");
                document.Save();
            }

            WordPackageSigningResult result = WordDocument.TrySignPackage(filePath, "ABCDZ123");

            Assert.False(result.Succeeded);
            Assert.Contains(result.Details, detail => detail.Contains("invalid character", System.StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Test_DigitalSignature_SignPackageCreatesCryptographicallyValidSignatureCrossPlatform() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureSignedByAdapter.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Package signing adapter proof");
                document.Save();
            }
            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true)) {
                package.ExtendedFilePropertiesPart?.Properties?.DigitalSignature?.Remove();
                package.ExtendedFilePropertiesPart?.Properties?.Save();
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
            Assert.Matches("^/_xmlsignatures/sig[0-9]*\\.xml$", result.SignaturePartUri!);
            Assert.NotNull(result.ValidationReport);
            Assert.True(result.ValidationReport!.IsStructurallyValid);
            Assert.Equal(WordSignatureValidationState.Passed, result.ValidationReport.CryptographicStatus);
            Assert.Equal(WordSignatureValidationState.Passed, result.ValidationReport.SignedPartDigestStatus);

            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, false)) {
                Assert.NotNull(package.DigitalSignatureOriginPart);
                Assert.NotEmpty(package.DigitalSignatureOriginPart!.XmlSignatureParts);
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                var validationOptions = new WordSignatureValidationOptions();
                validationOptions.CertificateValidation.DisableCertificateDownloads = false;
                validationOptions.CertificateValidation.ChainEvaluator = static (_, _) => true;
                WordSignatureValidationReport validation = document.ValidateSignatures(validationOptions);

                Assert.True(validation.HasSignatures);
                Assert.True(validation.IsStructurallyValid, string.Join(System.Environment.NewLine, validation.Findings));
                Assert.Equal(WordSignatureValidationState.Passed, validation.PackageStructureStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.XmlSignatureStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.SignedPartCoverageStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.CryptographicStatus);
                Assert.Equal(WordSignatureValidationState.Passed, validation.CertificateChainStatus);
                Assert.True(validation.IsValidUnderPolicy, string.Join(System.Environment.NewLine, validation.Findings));
                Assert.True(validation.SignatureInfo.SignatureParts.Count > 0);
                Assert.Contains(validation.SignatureInfo.SignatureParts.SelectMany(part => part.SignedReferences), reference => reference.HasDigestValue);
            }
        }

        [Fact]
        public void Test_DigitalSignature_SignPackagePreservesAConcurrentSourceChange() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureConcurrentSource.docx");
            string replacementPath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureConcurrentReplacement.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Original package state");
                document.Save();
            }
            using (WordDocument replacement = WordDocument.Create(replacementPath)) {
                replacement.AddParagraph("Concurrent package state");
                replacement.Save();
            }
            byte[] concurrentBytes = File.ReadAllBytes(replacementPath);

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            var options = new OfficePackageSigningOptions {
                BeforeCommit = (_, target) => File.WriteAllBytes(target, concurrentBytes)
            };

            OfficePackageSigningResult result = OfficePackageSignatureWriter.Sign(filePath, certificate, options);

            Assert.False(result.Succeeded);
            Assert.Contains(result.Details, detail =>
                detail.Contains("changed while its signature was being created", StringComparison.Ordinal));
            Assert.Equal(concurrentBytes, File.ReadAllBytes(filePath));
            using WordprocessingDocument preserved = WordprocessingDocument.Open(filePath, false);
            Assert.Null(preserved.DigitalSignatureOriginPart);
        }

        [Fact]
        public void Test_DigitalSignature_ValidateSignaturesUsesCurrentInMemoryDocumentState() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureLiveStateValidation.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed source state");
                document.Save();
            }

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordDocument.SignPackage(filePath, certificate);

            using WordDocument loaded = WordDocument.Load(filePath);
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;
            Assert.Equal(WordSignatureValidationState.Passed, loaded.ValidateSignatures(options).SignedPartDigestStatus);

            loaded.AddParagraph("Unsaved mutation");
            WordSignatureValidationReport validation = loaded.ValidateSignatures(options);

            Assert.Equal(WordSignatureValidationState.Failed, validation.SignedPartDigestStatus);
            Assert.False(Assert.Single(validation.Signatures).IsValidUnderPolicy);
            Assert.False(validation.IsValidUnderPolicy);
        }

        [Fact]
        public void Test_DigitalSignature_RejectsTamperedSignedPackageObject() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureTamperedSignedObject.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed-object digest validation");
                document.Save();
            }

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordDocument.SignPackage(filePath, certificate);
            TamperSignedPackageObject(filePath);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;

            WordSignatureValidationReport validation = loaded.ValidateSignatures(options);

            Assert.Equal(WordSignatureValidationState.Failed, validation.CryptographicStatus);
            Assert.Equal(WordSignatureValidationState.Passed, validation.SignedPartDigestStatus);
            Assert.False(validation.IsValidUnderPolicy);
            Assert.Contains(validation.Diagnostics, finding => finding.Code == "XmlSignatureInvalid");
        }

        [Fact]
        public void Test_DigitalSignature_RejectsUnsupportedSignedInfoTransformBeforeCryptographicValidation() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureUnsupportedSignedInfoTransform.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Unsupported SignedInfo transform");
                document.Save();
            }

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordDocument.SignPackage(filePath, certificate);
            AddUnsupportedSignedInfoTransform(filePath);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            WordSignatureValidationReport validation = loaded.ValidateSignatures();

            Assert.Equal(WordSignatureValidationState.Unsupported, validation.CryptographicStatus);
            Assert.Contains(validation.Diagnostics, finding => finding.Code == "UnsupportedSignedInfoTransform");
        }

        [Fact]
        public void Test_DigitalSignature_RejectsUnsupportedSignedInfoCanonicalizationBeforeCryptographicValidation() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureUnsupportedSignedInfoCanonicalization.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Unsupported SignedInfo canonicalization");
                document.Save();
            }

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordDocument.SignPackage(filePath, certificate);
            SetSignedInfoCanonicalizationMethod(filePath, SignedXml.XmlDsigXsltTransformUrl);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            WordSignatureValidationReport validation = loaded.ValidateSignatures();

            Assert.Equal(WordSignatureValidationState.Unsupported, validation.CryptographicStatus);
            Assert.Contains(validation.Diagnostics, finding => finding.Code == "UnsupportedSignedInfoCanonicalizationMethod");
        }

        [Fact]
        public void Test_DigitalSignature_CountersigningPreservesExistingSignedApplicationProperties() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "DigitalSignatures", "signed-valid.docx"));
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureCountersignPreservesApplicationProperties.docx");
            File.Copy(sourcePath, filePath, overwrite: true);

            string existingSignatureUri;
            using (WordDocument original = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            })) {
                existingSignatureUri = Assert.Single(original.InspectSignatures().SignatureParts).Uri;
            }

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult signing = WordDocument.SignPackage(filePath, certificate);

            Assert.True(signing.Succeeded);
            Assert.Equal(2, signing.SignatureCount);
            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, false)) {
                Assert.Null(package.ExtendedFilePropertiesPart?.Properties?.DigitalSignature);
            }
            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;
            WordSignatureValidationReport validation = loaded.ValidateSignatures(options);

            WordSignaturePartValidationResult existing = Assert.Single(validation.Signatures, signature =>
                string.Equals(signature.SignaturePart.Uri, existingSignatureUri, System.StringComparison.OrdinalIgnoreCase));
            Assert.Equal(WordSignatureValidationState.Passed, existing.CryptographicStatus);
            Assert.All(existing.SignaturePart.SignedReferences.Where(reference => reference.IsPackagePartReference), reference =>
                Assert.Equal(WordSignatureValidationState.Passed, reference.DigestVerificationStatus));
            Assert.Equal(WordSignatureValidationState.Passed, validation.SignedPartDigestStatus);
        }

        [Fact]
        public void Test_DigitalSignature_SignPackageValidatesCreatedSignatureIndependentlyOfExistingInvalidSignature() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureCountersignedInvalidSignature.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Countersign a package containing an invalid signature");
                document.Save();
            }
            AddDigitalSignatureMetadata(filePath, CreateSignatureXml(digestValue: "T2ZmaWNlSU1P"));

            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult result = WordDocument.SignPackage(filePath, certificate);

            Assert.True(result.Succeeded);
            Assert.Equal(2, result.SignatureCount);
            Assert.NotNull(result.ValidationReport);
            Assert.NotEqual(WordSignatureValidationState.Passed, result.ValidationReport!.CryptographicStatus);
            WordSignaturePartValidationResult created = Assert.IsType<WordSignaturePartValidationResult>(result.CreatedSignatureValidation);
            Assert.Equal(result.SignaturePartUri, created.SignaturePart.Uri, ignoreCase: true);
            Assert.Equal(WordSignatureValidationState.Passed, created.CryptographicStatus);
            Assert.NotEmpty(created.SignaturePart.SignedReferences);
            Assert.All(created.SignaturePart.SignedReferences, reference => {
                Assert.True(reference.IsPackagePartReference);
                Assert.True(reference.TargetPartExists);
                Assert.Equal(WordSignatureValidationState.Passed, reference.DigestVerificationStatus);
            });
        }

        [Fact]
        public void Test_DigitalSignature_ValidatesEmbeddedRfc3161TimestampAuthorityToken() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureTimestamped.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Timestamped OPC package signature");
                document.Save();
            }

            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            WordDocument.SignPackage(filePath, signer, new WordPackageSigningOptions {
                SignatureId = "OfficeIMOTimestampedSignature"
            });
            AddRfc3161Timestamp(
                filePath,
                timestampCorrectSignatureValue: true,
                canonicalizationAlgorithm: SignedXml.XmlDsigExcC14NTransformUrl);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.DisableCertificateDownloads = false;
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;
            options.TimestampCertificateValidation.DisableCertificateDownloads = false;
            options.TimestampCertificateValidation.ChainEvaluator = static (_, _) => true;

            WordSignatureValidationReport validation = loaded.ValidateSignatures(options);

            WordSignaturePartValidationResult signature = Assert.Single(validation.Signatures);
            Assert.Equal(WordSignatureValidationState.Passed, signature.CryptographicStatus);
            Assert.Equal(WordSignatureValidationState.Passed, signature.TimestampStatus);
            Assert.Equal(OfficeIMO.Security.SecurityValidationStatus.Valid, Assert.Single(signature.TimestampTokens).Status);
            Assert.True(validation.IsValidUnderPolicy, string.Join(System.Environment.NewLine, validation.Findings));
        }

        [Theory]
        [InlineData(SignedXml.XmlDsigExcC14NTransformUrl, "WithoutComments")]
        [InlineData(SignedXml.XmlDsigExcC14NWithCommentsTransformUrl, "WithComments")]
        public void Test_DigitalSignature_HonorsExclusiveCanonicalizationInclusiveNamespacesForTimestamp(
            string canonicalizationAlgorithm,
            string fileSuffix) {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureTimestampInclusiveNamespaces" + fileSuffix + ".docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Timestamped OPC package signature with an inclusive prefix list");
                document.Save();
            }

            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            WordDocument.SignPackage(filePath, signer, new WordPackageSigningOptions {
                SignatureId = "OfficeIMOTimestampedInclusiveNamespacesSignature"
            });
            AddRfc3161Timestamp(
                filePath,
                timestampCorrectSignatureValue: true,
                canonicalizationAlgorithm: canonicalizationAlgorithm,
                inclusiveNamespacesPrefixList: "proof");

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.DisableCertificateDownloads = false;
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;
            options.TimestampCertificateValidation.DisableCertificateDownloads = false;
            options.TimestampCertificateValidation.ChainEvaluator = static (_, _) => true;

            WordSignatureValidationReport validation = loaded.ValidateSignatures(options);

            WordSignaturePartValidationResult signature = Assert.Single(validation.Signatures);
            Assert.Equal(WordSignatureValidationState.Passed, signature.TimestampStatus);
            Assert.Equal(OfficeIMO.Security.SecurityValidationStatus.Valid, Assert.Single(signature.TimestampTokens).Status);
            Assert.True(signature.IsValidUnderPolicy, string.Join(System.Environment.NewLine, signature.Findings));
        }

        [Fact]
        public void Test_DigitalSignature_HonorsInheritedXmlAttributesForInclusiveTimestampCanonicalization() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureTimestampInheritedXmlLanguage.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Timestamped OPC signature with inherited XML context");
                document.Save();
            }

            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            WordDocument.SignPackage(filePath, signer, new WordPackageSigningOptions {
                SignatureId = "OfficeIMOTimestampedInheritedXmlLanguageSignature"
            });
            AddRfc3161Timestamp(
                filePath,
                timestampCorrectSignatureValue: true,
                canonicalizationAlgorithm: SignedXml.XmlDsigC14NTransformUrl,
                inheritedXmlLanguage: "en-GB");

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.DisableCertificateDownloads = false;
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;
            options.TimestampCertificateValidation.DisableCertificateDownloads = false;
            options.TimestampCertificateValidation.ChainEvaluator = static (_, _) => true;

            WordSignaturePartValidationResult signature = Assert.Single(loaded.ValidateSignatures(options).Signatures);

            Assert.Equal(WordSignatureValidationState.Passed, signature.TimestampStatus);
            Assert.Equal(OfficeIMO.Security.SecurityValidationStatus.Valid, Assert.Single(signature.TimestampTokens).Status);
        }

        [Fact]
        public void Test_DigitalSignature_RejectsTimestampWithMismatchedMessageImprint() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureBadTimestamp.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Mismatched RFC 3161 timestamp imprint");
                document.Save();
            }

            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            WordDocument.SignPackage(filePath, signer);
            AddRfc3161Timestamp(filePath, timestampCorrectSignatureValue: false);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;
            options.TimestampCertificateValidation.ChainEvaluator = static (_, _) => true;

            WordSignatureValidationReport validation = loaded.ValidateSignatures(options);

            Assert.Equal(WordSignatureValidationState.Passed, validation.CryptographicStatus);
            Assert.Equal(WordSignatureValidationState.Failed, validation.TimestampStatus);
            Assert.False(validation.IsValidUnderPolicy);
            Assert.Contains(validation.Diagnostics, finding => finding.Code == "TimestampImprintMismatch");
        }

        [Fact]
        public void Test_DigitalSignature_RejectsMalformedEmbeddedTimestampToken() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureMalformedTimestamp.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Malformed RFC 3161 timestamp token");
                document.Save();
            }

            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            WordDocument.SignPackage(filePath, signer);
            AddRfc3161Timestamp(filePath, timestampCorrectSignatureValue: true, timestampTokenText: "not-valid-base64!");

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;

            WordSignatureValidationReport validation = loaded.ValidateSignatures(options);

            Assert.Equal(WordSignatureValidationState.Failed, validation.TimestampStatus);
            Assert.False(validation.IsValidUnderPolicy);
            Assert.Contains(validation.Diagnostics, finding => finding.Code == "TimestampMalformed");
        }

        [Fact]
        public void Test_DigitalSignature_RejectsOversizedTimestampBeforeDecoding() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureOversizedTimestamp.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Oversized RFC 3161 timestamp token");
                document.Save();
            }

            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            WordDocument.SignPackage(filePath, signer);
            AddRfc3161Timestamp(filePath, timestampCorrectSignatureValue: true, timestampTokenText: new string('A', 512));

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            var options = new WordSignatureValidationOptions { MaxTimestampBytes = 32 };
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;

            WordSignatureValidationReport validation = loaded.ValidateSignatures(options);

            Assert.False(validation.IsValidUnderPolicy);
            Assert.Equal(WordSignatureValidationState.Passed, validation.CryptographicStatus);
            Assert.Equal(WordSignatureValidationState.Failed, validation.TimestampStatus);
            Assert.Contains(validation.Diagnostics, finding => finding.Code == "TimestampResourceLimitExceeded");
        }

        [Fact]
        public void Test_DigitalSignature_IgnoresTimestampLikeElementsOutsideXades() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureUnrelatedTimestampElement.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Unrelated timestamp-like signed object");
                document.Save();
            }

            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            WordDocument.SignPackage(filePath, signer);
            AddUnrelatedTimestampLikeObject(filePath);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;

            WordSignatureValidationReport validation = loaded.ValidateSignatures(options);

            Assert.Equal(WordSignatureValidationState.Passed, validation.CryptographicStatus);
            Assert.Equal(WordSignatureValidationState.NotPresent, validation.TimestampStatus);
            Assert.DoesNotContain(validation.Diagnostics, finding => finding.Code == "TimestampMalformed");
        }

        [Fact]
        public void Test_DigitalSignature_IgnoresXadesTimestampForAnotherSignatureTarget() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureUnrelatedXadesTimestamp.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Unrelated XAdES timestamp object");
                document.Save();
            }

            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            WordDocument.SignPackage(filePath, signer);
            AddUnrelatedXadesTimestampObject(filePath);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;

            WordSignatureValidationReport validation = loaded.ValidateSignatures(options);

            Assert.Equal(WordSignatureValidationState.Passed, validation.CryptographicStatus);
            Assert.Equal(WordSignatureValidationState.NotPresent, validation.TimestampStatus);
            Assert.DoesNotContain(validation.Diagnostics, finding => finding.Code == "TimestampMalformed");
        }

        [Fact]
        public void Test_DigitalSignature_SelectsTrustedCertificateFromEveryMatchingPublicKeyCandidate() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureMatchingCertificateCandidates.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Matching signer certificate candidates");
                document.Save();
            }

            using RSA signingKey = RSA.Create(2048);
            using X509Certificate2 rejectedCertificate = CreateSelfSignedSigningCertificate(
                signingKey,
                "CN=OfficeIMO Rejected Matching Signer");
            using X509Certificate2 acceptedCertificate = CreateSelfSignedSigningCertificate(
                signingKey,
                "CN=OfficeIMO Accepted Matching Signer");
            WordDocument.SignPackage(filePath, rejectedCertificate);
            AppendEmbeddedSignerCertificate(filePath, acceptedCertificate);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            var evaluatedSubjects = new List<string>();
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.DisableCertificateDownloads = false;
            options.CertificateValidation.ChainEvaluator = (certificate, _) => {
                evaluatedSubjects.Add(certificate.Subject);
                return certificate.Subject.Contains("Accepted Matching Signer", StringComparison.Ordinal);
            };

            WordSignatureValidationReport validation = loaded.ValidateSignatures(options);

            Assert.Equal(WordSignatureValidationState.Passed, validation.CryptographicStatus);
            Assert.Equal(WordSignatureValidationState.Passed, validation.CertificateChainStatus);
            Assert.True(validation.IsValidUnderPolicy, string.Join(Environment.NewLine, validation.Findings));
            Assert.Contains(evaluatedSubjects, subject => subject.Contains("Rejected Matching Signer", StringComparison.Ordinal));
            Assert.Contains(evaluatedSubjects, subject => subject.Contains("Accepted Matching Signer", StringComparison.Ordinal));
        }

        [Fact]
        public void Test_DigitalSignature_ReportsUnsupportedTimestampCanonicalizationWithoutFailure() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureUnsupportedTimestampCanonicalization.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Unsupported timestamp canonicalization");
                document.Save();
            }

            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            WordDocument.SignPackage(filePath, signer);
            AddRfc3161Timestamp(
                filePath,
                timestampCorrectSignatureValue: true,
                canonicalizationAlgorithm: "urn:officeimo:unsupported-canonicalization");

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;

            WordSignatureValidationReport validation = loaded.ValidateSignatures(options);

            Assert.Equal(WordSignatureValidationState.Passed, validation.CryptographicStatus);
            Assert.Equal(WordSignatureValidationState.Unsupported, validation.TimestampStatus);
            Assert.Contains(validation.Diagnostics, finding => finding.Code == "TimestampCanonicalizationUnsupported");
            Assert.DoesNotContain(validation.Diagnostics, finding => finding.Code == "TimestampValidationFailed");
        }

        [Fact]
        public void Test_DigitalSignature_TrySignPackageFailsWhenRequestedPartIsMissing() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureMissingRequestedPart.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Package signing missing requested part proof");
                document.Save();
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
                document.Save();
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

#if NET472
        [Fact]
        public void Test_DigitalSignature_SignPackageCanResolveCertificateFromStoreOnSupportedAdapter() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureSignedByStoreCertificate.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Package signing certificate-store proof");
                document.Save();
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
#endif

        [Fact]
        public void Test_DigitalSignature_TrySignPackageReportsMissingStoreCertificate() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureMissingStoreCertificate.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Missing store certificate proof");
                document.Save();
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

        [Fact]
        public void Test_DigitalSignature_TrySignPackageReportsDisposedPrivateKeyProviderFailure() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureDisposedCertificate.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Unavailable private-key provider");
                document.Save();
            }
            X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            certificate.Dispose();

            WordPackageSigningResult result = WordDocument.TrySignPackage(filePath, certificate);

            Assert.False(result.Succeeded);
            Assert.Contains(result.Details, detail =>
                detail.Contains("private key", System.StringComparison.OrdinalIgnoreCase) ||
                detail.Contains("signing failed", System.StringComparison.OrdinalIgnoreCase));
        }

        private static byte[] CreateSignatureXml(
            string referenceUri = "/word/document.xml?ContentType=application%2Fvnd.openxmlformats-officedocument.wordprocessingml.document.main%2Bxml",
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

        private static byte[] CreateSignatureXmlWithCertificates(params string[] certificates) {
            return Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\">" +
                "<SignedInfo><SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" /></SignedInfo>" +
                "<KeyInfo><X509Data>" +
                string.Concat(certificates.Select(certificate => "<X509Certificate>" + certificate + "</X509Certificate>")) +
                "</X509Data></KeyInfo><SignatureValue>AA==</SignatureValue></Signature>");
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

        private static void AddDigitalSignatureMetadata(string filePath, byte[] signatureBytes, int signatureCount = 1) {
            using WordprocessingDocument package = WordprocessingDocument.Open(filePath, true);
            package.AddDigitalSignatureOriginPart();
            for (int i = 0; i < signatureCount; i++) {
                XmlSignaturePart signaturePart = package.DigitalSignatureOriginPart!.AddNewPart<XmlSignaturePart>();
                using var stream = new MemoryStream(signatureBytes);
                signaturePart.FeedData(stream);
            }

            ExtendedFilePropertiesPart appPart = package.ExtendedFilePropertiesPart ?? package.AddExtendedFilePropertiesPart();
            appPart.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
            appPart.Properties.DigitalSignature = new DigitalSignature();
            appPart.Properties.Save();
        }

        private static X509Certificate2 CreateSelfSignedSigningCertificate(string subjectName = "CN=OfficeIMO Package Signing Test") {
            using RSA rsa = RSA.Create(2048);
            return CreateSelfSignedSigningCertificate(rsa, subjectName);
        }

        private static X509Certificate2 CreateSelfSignedSigningCertificate(RSA rsa, string subjectName) {
            var request = new CertificateRequest(
                subjectName,
                rsa,
                HashAlgorithmName.SHA256,
                RSASignaturePadding.Pkcs1);

            request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: false));

            using X509Certificate2 certificate = request.CreateSelfSigned(
                DateTimeOffset.UtcNow.AddDays(-1),
                DateTimeOffset.UtcNow.AddDays(1));

            return new X509Certificate2(certificate.Export(X509ContentType.Pfx), (string?)null, X509KeyStorageFlags.Exportable);
        }

        private static void AddRfc3161Timestamp(
            string filePath,
            bool timestampCorrectSignatureValue,
            string? timestampTokenText = null,
            string? canonicalizationAlgorithm = null,
            string? inclusiveNamespacesPrefixList = null,
            string? inheritedXmlLanguage = null) {
            using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry signatureEntry = archive.Entries.Single(entry =>
                entry.FullName.Contains("_xmlsignatures", System.StringComparison.OrdinalIgnoreCase) &&
                entry.FullName.EndsWith(".xml", System.StringComparison.OrdinalIgnoreCase) &&
                !entry.FullName.Contains("_rels", System.StringComparison.OrdinalIgnoreCase));
            var signatureXml = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            using (Stream source = signatureEntry.Open()) {
                signatureXml.Load(source);
            }

            var namespaceManager = new XmlNamespaceManager(signatureXml.NameTable);
            namespaceManager.AddNamespace("ds", "http://www.w3.org/2000/09/xmldsig#");
            XmlElement signature = signatureXml.DocumentElement!;
            XmlElement signatureValue = (XmlElement)signature.SelectSingleNode("ds:SignatureValue", namespaceManager)!;
            if (!string.IsNullOrWhiteSpace(inheritedXmlLanguage)) {
                XmlAttribute xmlLanguage = signatureXml.CreateAttribute("xml", "lang", "http://www.w3.org/XML/1998/namespace");
                xmlLanguage.Value = inheritedXmlLanguage!;
                signature.Attributes.Append(xmlLanguage);
            }
            if (!string.IsNullOrWhiteSpace(inclusiveNamespacesPrefixList)) {
                signatureValue.SetAttribute("xmlns:proof", "urn:officeimo:timestamp-proof");
            }
            byte[] timestampedValue = timestampCorrectSignatureValue
                ? CanonicalizeSignatureValue(signatureValue, canonicalizationAlgorithm, inclusiveNamespacesPrefixList)
                : Encoding.UTF8.GetBytes("not the XML signature value");
            byte[] timestampToken = CreateRfc3161TimestampToken(timestampedValue);

            XmlElement dataObject = signatureXml.CreateElement("ds", "Object", "http://www.w3.org/2000/09/xmldsig#");
            XmlElement qualifyingProperties = signatureXml.CreateElement("xades", "QualifyingProperties", "http://uri.etsi.org/01903/v1.3.2#");
            qualifyingProperties.SetAttribute("Target", "#" + signature.GetAttribute("Id"));
            XmlElement unsignedProperties = signatureXml.CreateElement("xades", "UnsignedProperties", qualifyingProperties.NamespaceURI);
            XmlElement unsignedSignatureProperties = signatureXml.CreateElement("xades", "UnsignedSignatureProperties", qualifyingProperties.NamespaceURI);
            XmlElement signatureTimeStamp = signatureXml.CreateElement("xades", "SignatureTimeStamp", qualifyingProperties.NamespaceURI);
            if (!string.IsNullOrWhiteSpace(canonicalizationAlgorithm)) {
                XmlElement canonicalizationMethod = signatureXml.CreateElement("ds", "CanonicalizationMethod", SignedXml.XmlDsigNamespaceUrl);
                canonicalizationMethod.SetAttribute("Algorithm", canonicalizationAlgorithm);
                if (!string.IsNullOrWhiteSpace(inclusiveNamespacesPrefixList)) {
                    XmlElement inclusiveNamespaces = signatureXml.CreateElement(
                        "ec",
                        "InclusiveNamespaces",
                        SignedXml.XmlDsigExcC14NTransformUrl);
                    inclusiveNamespaces.SetAttribute("PrefixList", inclusiveNamespacesPrefixList);
                    canonicalizationMethod.AppendChild(inclusiveNamespaces);
                }
                signatureTimeStamp.AppendChild(canonicalizationMethod);
            }
            XmlElement encapsulated = signatureXml.CreateElement("xades", "EncapsulatedTimeStamp", qualifyingProperties.NamespaceURI);
            encapsulated.InnerText = timestampTokenText ?? System.Convert.ToBase64String(timestampToken);
            signatureTimeStamp.AppendChild(encapsulated);
            unsignedSignatureProperties.AppendChild(signatureTimeStamp);
            unsignedProperties.AppendChild(unsignedSignatureProperties);
            qualifyingProperties.AppendChild(unsignedProperties);
            dataObject.AppendChild(qualifyingProperties);
            signature.AppendChild(dataObject);

            using Stream destination = signatureEntry.Open();
            destination.SetLength(0);
            signatureXml.Save(destination);
        }

        private static void AddUnrelatedTimestampLikeObject(string filePath) {
            using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry signatureEntry = archive.Entries.Single(entry =>
                entry.FullName.Contains("_xmlsignatures", System.StringComparison.OrdinalIgnoreCase) &&
                entry.FullName.EndsWith(".xml", System.StringComparison.OrdinalIgnoreCase) &&
                !entry.FullName.Contains("_rels", System.StringComparison.OrdinalIgnoreCase));
            var signatureXml = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            using (Stream source = signatureEntry.Open()) {
                signatureXml.Load(source);
            }

            XmlElement dataObject = signatureXml.CreateElement("ds", "Object", SignedXml.XmlDsigNamespaceUrl);
            XmlElement unrelated = signatureXml.CreateElement("custom", "EncapsulatedTimeStamp", "urn:officeimo:business-data");
            unrelated.InnerText = "not-an-rfc3161-token";
            dataObject.AppendChild(unrelated);
            signatureXml.DocumentElement!.AppendChild(dataObject);

            using Stream destination = signatureEntry.Open();
            destination.SetLength(0);
            signatureXml.Save(destination);
        }

        private static void AddUnrelatedXadesTimestampObject(string filePath) {
            using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry signatureEntry = archive.Entries.Single(entry =>
                entry.FullName.Contains("_xmlsignatures", StringComparison.OrdinalIgnoreCase) &&
                entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                !entry.FullName.Contains("_rels", StringComparison.OrdinalIgnoreCase));
            var signatureXml = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            using (Stream source = signatureEntry.Open()) {
                signatureXml.Load(source);
            }

            const string xadesNamespace = "http://uri.etsi.org/01903/v1.3.2#";
            XmlElement dataObject = signatureXml.CreateElement("ds", "Object", SignedXml.XmlDsigNamespaceUrl);
            XmlElement qualifyingProperties = signatureXml.CreateElement("xades", "QualifyingProperties", xadesNamespace);
            qualifyingProperties.SetAttribute("Target", "#UnrelatedNestedSignature");
            XmlElement unsignedProperties = signatureXml.CreateElement("xades", "UnsignedProperties", xadesNamespace);
            XmlElement unsignedSignatureProperties = signatureXml.CreateElement("xades", "UnsignedSignatureProperties", xadesNamespace);
            XmlElement signatureTimeStamp = signatureXml.CreateElement("xades", "SignatureTimeStamp", xadesNamespace);
            XmlElement token = signatureXml.CreateElement("xades", "EncapsulatedTimeStamp", xadesNamespace);
            token.InnerText = "not-valid-base64!";
            signatureTimeStamp.AppendChild(token);
            unsignedSignatureProperties.AppendChild(signatureTimeStamp);
            unsignedProperties.AppendChild(unsignedSignatureProperties);
            qualifyingProperties.AppendChild(unsignedProperties);
            dataObject.AppendChild(qualifyingProperties);
            signatureXml.DocumentElement!.AppendChild(dataObject);

            using Stream destination = signatureEntry.Open();
            destination.SetLength(0);
            signatureXml.Save(destination);
        }

        private static void AppendEmbeddedSignerCertificate(string filePath, X509Certificate2 certificate) {
            using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry signatureEntry = archive.Entries.Single(entry =>
                entry.FullName.Contains("_xmlsignatures", StringComparison.OrdinalIgnoreCase) &&
                entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                !entry.FullName.Contains("_rels", StringComparison.OrdinalIgnoreCase));
            var signatureXml = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            using (Stream source = signatureEntry.Open()) {
                signatureXml.Load(source);
            }

            var namespaceManager = new XmlNamespaceManager(signatureXml.NameTable);
            namespaceManager.AddNamespace("ds", SignedXml.XmlDsigNamespaceUrl);
            XmlElement x509Data = (XmlElement)signatureXml.SelectSingleNode("/ds:Signature/ds:KeyInfo/ds:X509Data", namespaceManager)!;
            XmlElement encodedCertificate = signatureXml.CreateElement("ds", "X509Certificate", SignedXml.XmlDsigNamespaceUrl);
            encodedCertificate.InnerText = Convert.ToBase64String(certificate.Export(X509ContentType.Cert));
            x509Data.AppendChild(encodedCertificate);

            using Stream destination = signatureEntry.Open();
            destination.SetLength(0);
            signatureXml.Save(destination);
        }

        private static void TamperSignedPackageObject(string filePath) {
            using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry signatureEntry = archive.Entries.Single(entry =>
                entry.FullName.Contains("_xmlsignatures", System.StringComparison.OrdinalIgnoreCase) &&
                entry.FullName.EndsWith(".xml", System.StringComparison.OrdinalIgnoreCase) &&
                !entry.FullName.Contains("_rels", System.StringComparison.OrdinalIgnoreCase));
            var signatureXml = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            using (Stream source = signatureEntry.Open()) {
                signatureXml.Load(source);
            }

            var namespaceManager = new XmlNamespaceManager(signatureXml.NameTable);
            namespaceManager.AddNamespace("ds", SignedXml.XmlDsigNamespaceUrl);
            XmlElement manifest = (XmlElement)signatureXml.SelectSingleNode("/ds:Signature/ds:Object/ds:Manifest", namespaceManager)!;
            manifest.SetAttribute("tampered", "true");

            using Stream destination = signatureEntry.Open();
            destination.SetLength(0);
            signatureXml.Save(destination);
        }

        private static void AddUnsupportedSignedInfoTransform(string filePath) {
            using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry signatureEntry = archive.Entries.Single(entry =>
                entry.FullName.Contains("_xmlsignatures", System.StringComparison.OrdinalIgnoreCase) &&
                entry.FullName.EndsWith(".xml", System.StringComparison.OrdinalIgnoreCase) &&
                !entry.FullName.Contains("_rels", System.StringComparison.OrdinalIgnoreCase));
            var signatureXml = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            using (Stream source = signatureEntry.Open()) {
                signatureXml.Load(source);
            }

            var namespaceManager = new XmlNamespaceManager(signatureXml.NameTable);
            namespaceManager.AddNamespace("ds", SignedXml.XmlDsigNamespaceUrl);
            XmlElement reference = (XmlElement)signatureXml.SelectSingleNode("/ds:Signature/ds:SignedInfo/ds:Reference", namespaceManager)!;
            XmlElement transforms = signatureXml.CreateElement("ds", "Transforms", SignedXml.XmlDsigNamespaceUrl);
            XmlElement transform = signatureXml.CreateElement("ds", "Transform", SignedXml.XmlDsigNamespaceUrl);
            transform.SetAttribute("Algorithm", SignedXml.XmlDsigXsltTransformUrl);
            transforms.AppendChild(transform);
            reference.PrependChild(transforms);

            using Stream destination = signatureEntry.Open();
            destination.SetLength(0);
            signatureXml.Save(destination);
        }

        private static void SetSignedInfoCanonicalizationMethod(string filePath, string algorithm) {
            using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry signatureEntry = archive.Entries.Single(entry =>
                entry.FullName.Contains("_xmlsignatures", System.StringComparison.OrdinalIgnoreCase) &&
                entry.FullName.EndsWith(".xml", System.StringComparison.OrdinalIgnoreCase) &&
                !entry.FullName.Contains("_rels", System.StringComparison.OrdinalIgnoreCase));
            var signatureXml = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            using (Stream source = signatureEntry.Open()) {
                signatureXml.Load(source);
            }

            var namespaceManager = new XmlNamespaceManager(signatureXml.NameTable);
            namespaceManager.AddNamespace("ds", SignedXml.XmlDsigNamespaceUrl);
            XmlElement canonicalization = (XmlElement)signatureXml.SelectSingleNode(
                "/ds:Signature/ds:SignedInfo/ds:CanonicalizationMethod",
                namespaceManager)!;
            canonicalization.SetAttribute("Algorithm", algorithm);

            using Stream destination = signatureEntry.Open();
            destination.SetLength(0);
            signatureXml.Save(destination);
        }

        private static byte[] CanonicalizeSignatureValue(
            XmlElement signatureValue,
            string? algorithm,
            string? inclusiveNamespacesPrefixList = null) {
            var input = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            XmlElement imported = (XmlElement)input.ImportNode(signatureValue, deep: true);
            var inheritedNames = new HashSet<string>(StringComparer.Ordinal);
            bool includeInheritedXmlAttributes =
                string.IsNullOrWhiteSpace(algorithm) ||
                algorithm == SignedXml.XmlDsigC14NTransformUrl ||
                algorithm == SignedXml.XmlDsigC14NWithCommentsTransformUrl;
            for (XmlElement? ancestor = signatureValue; ancestor != null; ancestor = ancestor.ParentNode as XmlElement) {
                foreach (XmlAttribute attribute in ancestor.Attributes) {
                    bool isNamespace = attribute.Prefix == "xmlns" || attribute.Name == "xmlns";
                    bool isInheritedXmlAttribute = includeInheritedXmlAttributes &&
                        attribute.NamespaceURI == "http://www.w3.org/XML/1998/namespace";
                    if (!isNamespace && !isInheritedXmlAttribute) continue;
                    string attributeKey = attribute.NamespaceURI + "\0" + attribute.LocalName;
                    if (!inheritedNames.Add(attributeKey) || imported.HasAttribute(attribute.LocalName, attribute.NamespaceURI)) continue;
                    imported.Attributes.Append((XmlAttribute)input.ImportNode(attribute, deep: true));
                }
            }
            input.AppendChild(imported);
            Transform transform;
            if (algorithm == SignedXml.XmlDsigExcC14NTransformUrl ||
                algorithm == SignedXml.XmlDsigExcC14NWithCommentsTransformUrl) {
                transform = new XmlDsigExcC14NTransform(
                    includeComments: algorithm == SignedXml.XmlDsigExcC14NWithCommentsTransformUrl) {
                    InclusiveNamespacesPrefixList = inclusiveNamespacesPrefixList
                };
            } else {
                transform = new XmlDsigC14NTransform();
            }
            transform.LoadInput(input);
            using Stream canonical = (Stream)transform.GetOutput(typeof(Stream));
            using var output = new MemoryStream();
            canonical.CopyTo(output);
            return output.ToArray();
        }

        private static byte[] CreateRfc3161TimestampToken(byte[] timestampedData) {
            using X509Certificate2 certificate = CreateTimestampCertificate();
            using RSA rsa = certificate.GetRSAPrivateKey() ?? throw new System.InvalidOperationException();
            Org.BouncyCastle.X509.X509Certificate bcCertificate =
                Org.BouncyCastle.Security.DotNetUtilities.FromX509Certificate(certificate);
            Org.BouncyCastle.Crypto.AsymmetricKeyParameter privateKey =
                Org.BouncyCastle.Security.DotNetUtilities.GetRsaKeyPair(rsa).Private;
            var signerFactory = new Org.BouncyCastle.Crypto.Operators.Asn1SignatureFactory("SHA256WITHRSA", privateKey);
            Org.BouncyCastle.Cms.SignerInfoGenerator signer =
                new Org.BouncyCastle.Cms.SignerInfoGeneratorBuilder().Build(signerFactory, bcCertificate);
            var generator = new Org.BouncyCastle.Tsp.TimeStampTokenGenerator(
                signer,
                Org.BouncyCastle.Crypto.Operators.Asn1DigestFactory.Get("SHA256"),
                new Org.BouncyCastle.Asn1.DerObjectIdentifier("1.3.6.1.4.1.59069.1.1"),
                isIssuerSerialIncluded: true);
            generator.SetCertificates(new SingleCertificateStore(bcCertificate));
            var requestGenerator = new Org.BouncyCastle.Tsp.TimeStampRequestGenerator();
            requestGenerator.SetCertReq(true);
            byte[] imprint = Org.BouncyCastle.Security.DigestUtilities.CalculateDigest("SHA256", timestampedData);
            Org.BouncyCastle.Tsp.TimeStampRequest request = requestGenerator.Generate(
                Org.BouncyCastle.Tsp.TspAlgorithms.Sha256,
                imprint);
            return generator.Generate(request, Org.BouncyCastle.Math.BigInteger.One, System.DateTime.UtcNow).GetEncoded();
        }

        private static X509Certificate2 CreateTimestampCertificate() {
            using RSA rsa = RSA.Create(2048);
            var request = new CertificateRequest(
                "CN=OfficeIMO Word Test TSA",
                rsa,
                HashAlgorithmName.SHA256,
                RSASignaturePadding.Pkcs1);
            request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
            request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(
                new OidCollection { new Oid("1.3.6.1.5.5.7.3.8") },
                critical: true));
            request.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(request.PublicKey, critical: false));
            return request.CreateSelfSigned(System.DateTimeOffset.UtcNow.AddMinutes(-5), System.DateTimeOffset.UtcNow.AddDays(1));
        }

        private sealed class SingleCertificateStore :
            Org.BouncyCastle.Utilities.Collections.IStore<Org.BouncyCastle.X509.X509Certificate> {
            private readonly Org.BouncyCastle.X509.X509Certificate _certificate;

            internal SingleCertificateStore(Org.BouncyCastle.X509.X509Certificate certificate) {
                _certificate = certificate;
            }

            public System.Collections.Generic.IEnumerable<Org.BouncyCastle.X509.X509Certificate> EnumerateMatches(
                Org.BouncyCastle.Utilities.Collections.ISelector<Org.BouncyCastle.X509.X509Certificate> selector) {
                if (selector == null || selector.Match(_certificate)) yield return _certificate;
            }
        }

        private static void AddCertificateToCurrentUserStore(X509Certificate2 certificate) {
            using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadWrite);
            store.Add(certificate);
        }

        private static bool? ReadDisableCertificateDownloads(X509ChainPolicy policy) {
            System.Reflection.PropertyInfo? property = typeof(X509ChainPolicy).GetProperty("DisableCertificateDownloads");
            object? value = property?.GetValue(policy, null);
            return value is bool result ? result : null;
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

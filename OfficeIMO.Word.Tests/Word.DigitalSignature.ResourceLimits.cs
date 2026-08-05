using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using System.IO.Compression;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_DigitalSignature_BoundsContentTypesBeforeParsing() {
            using var packageStream = new MemoryStream();
            using (var archive = new ZipArchive(packageStream, ZipArchiveMode.Create, leaveOpen: true)) {
                ZipArchiveEntry contentTypes = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
                using var writer = new StreamWriter(contentTypes.Open(), Encoding.UTF8);
                writer.Write("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
                writer.Write(new string(' ', 4096));
                writer.Write("</Types>");
            }

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                new OfficePackageSignatureArchive(packageStream.ToArray(), maxParts: 10, maxPartBytes: 1024));

            Assert.Contains("exceeds", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_DigitalSignature_SigningRejectsPackagePartLimitBeforeOpenXmlParsing() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignaturePreOpenPartLimit.docx");
            using (var archive = ZipFile.Open(filePath, ZipArchiveMode.Create)) {
                archive.CreateEntry("[Content_Types].xml");
                archive.CreateEntry("extra.bin");
            }
            byte[] originalBytes = File.ReadAllBytes(filePath);
            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();

            WordPackageSigningResult result = WordDocument.TrySignPackage(filePath, SecurityProvider,
                signer,
                new WordPackageSigningOptions { MaxPackageParts = 1 });

            Assert.False(result.Succeeded);
            Assert.Contains(result.Details, detail =>
                detail.Contains("more than 1 ZIP entries", StringComparison.OrdinalIgnoreCase));
            Assert.Equal(originalBytes, File.ReadAllBytes(filePath));
        }

        [Fact]
        public void Test_DigitalSignature_ReadWriteValidationRejectsEncodedPackageBeforeSnapshotCopy() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureReadWritePackageBudget.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph(new string('x', 4096));
                document.Save();
            }
            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult signing = WordDocument.SignPackage(filePath, SecurityProvider, signer);
            Assert.True(signing.Succeeded);
            long packageLength = new FileInfo(filePath).Length;

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadWrite,
                MaxInputBytes = packageLength
            });
            WordSignatureValidationReport validation = loaded.ValidateSignatures(SecurityProvider,
                new WordSignatureValidationOptions { MaxPackageBytes = packageLength - 1 });

            WordSignatureValidationFinding finding = Assert.Single(validation.Diagnostics,
                diagnostic => diagnostic.Code == "PackageByteLimitExceeded");
            Assert.Contains("exceeds", finding.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_DigitalSignature_DeduplicatesAdditionalCertificatesBeforeWriting() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureDuplicateCertificates.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Duplicate signing certificates");
                document.Save();
            }

            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            using X509Certificate2 additional = CreateSelfSignedSigningCertificate("CN=OfficeIMO Duplicate Additional Certificate");
            WordPackageSigningResult result = WordDocument.SignPackage(filePath, SecurityProvider,
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
        public void Test_DigitalSignature_UsesBoundedCallerCertificateAsCryptographicCandidate() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureCallerCertificate.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Caller-supplied signer certificate");
                document.Save();
            }
            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            WordPackageSigningResult signing = WordDocument.SignPackage(filePath, SecurityProvider, signer);
            Assert.True(signing.Succeeded);

            using (FileStream packageStream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite))
            using (var archive = new ZipArchive(packageStream, ZipArchiveMode.Update)) {
                ZipArchiveEntry signatureEntry = archive.Entries.Single(entry =>
                    entry.FullName.StartsWith("_xmlsignatures/sig", StringComparison.OrdinalIgnoreCase) &&
                    entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
                string entryName = signatureEntry.FullName;
                string signatureXml;
                using (var reader = new StreamReader(signatureEntry.Open(), Encoding.UTF8, true)) {
                    signatureXml = reader.ReadToEnd();
                }
                string withoutCertificate = System.Text.RegularExpressions.Regex.Replace(
                    signatureXml,
                    "<KeyInfo>.*?</KeyInfo>",
                    string.Empty,
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                Assert.NotEqual(signatureXml, withoutCertificate);
                signatureEntry.Delete();
                ZipArchiveEntry replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
                using Stream output = replacement.Open();
                byte[] bytes = Encoding.UTF8.GetBytes(withoutCertificate);
                output.Write(bytes, 0, bytes.Length);
            }

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            WordSignatureValidationReport missing = loaded.ValidateSignatures(SecurityProvider);
            using X509Certificate2 unrelated = CreateSelfSignedSigningCertificate(
                "CN=OfficeIMO Unrelated Caller Certificate");
            var boundedOptions = new WordSignatureValidationOptions { MaxCertificates = 1 };
            boundedOptions.CertificateValidation.ExtraCertificates.Add(unrelated);
            boundedOptions.CertificateValidation.ExtraCertificates.Add(signer);
            WordSignatureValidationReport bounded = loaded.ValidateSignatures(SecurityProvider, boundedOptions);
            var options = new WordSignatureValidationOptions();
            options.CertificateValidation.ValidateChain = false;
            options.CertificateValidation.ExtraCertificates.Add(signer);
            WordSignatureValidationReport supplied = loaded.ValidateSignatures(SecurityProvider, options);

            Assert.Contains(missing.Diagnostics, finding => finding.Code == "SignerCertificateMissing");
            Assert.Contains(bounded.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
            Assert.True(
                supplied.CryptographicStatus == WordSignatureValidationState.Passed,
                string.Join(" | ", supplied.Diagnostics.Select(finding =>
                    finding.Code + ": " + finding.Message)));
            Assert.DoesNotContain(supplied.Diagnostics, finding => finding.Code == "SignerCertificateMissing");
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
            WordPackageSigningResult result = WordDocument.TrySignPackage(filePath, SecurityProvider,
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

            using WordDocument loaded = WordDocument.Load(filePath);
            WordSignatureValidationReport bounded = loaded.ValidateSignatures(SecurityProvider, new WordSignatureValidationOptions {
                MaxTotalDigestBytes = 768
            });
            WordSignatureValidationReport allowed = loaded.ValidateSignatures(SecurityProvider, new WordSignatureValidationOptions {
                MaxTotalDigestBytes = 4096
            });

            Assert.Contains(bounded.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
            Assert.DoesNotContain(allowed.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
            Assert.Contains(allowed.Diagnostics, finding => finding.Code == "XmlSignatureInvalid");
        }

        [Fact]
        public void Test_DigitalSignature_BoundsLocalReferenceDigestWorkAcrossSignatureParts() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureLocalReferenceMultiPartBudget.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Local SignedInfo reference work across signatures");
                document.Save();
            }
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            string encodedCertificate = Convert.ToBase64String(certificate.Export(X509ContentType.Cert));
            string digest = "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>T2ZmaWNlSU1P</DigestValue>";
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\">" +
                "<SignedInfo><CanonicalizationMethod Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\" />" +
                "<SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                "<Reference URI=\"#payload\">" + digest + "</Reference></SignedInfo>" +
                "<SignatureValue>AA==</SignatureValue>" +
                "<KeyInfo><X509Data><X509Certificate>" + encodedCertificate + "</X509Certificate></X509Data></KeyInfo>" +
                "<Object Id=\"payload\">" + new string('x', 512) + "</Object></Signature>");
            AddDigitalSignatureMetadata(filePath, signatureBytes, signatureCount: 2);

            using WordDocument loaded = WordDocument.Load(filePath);
            WordSignatureValidationReport bounded = loaded.ValidateSignatures(SecurityProvider, new WordSignatureValidationOptions {
                MaxTotalDigestBytes = 768
            });
            WordSignatureValidationReport allowed = loaded.ValidateSignatures(SecurityProvider, new WordSignatureValidationOptions {
                MaxTotalDigestBytes = 4096
            });

            Assert.Contains(bounded.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
            Assert.DoesNotContain(allowed.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
        }

        [Fact]
        public void Test_DigitalSignature_SharesDigestWorkAcrossInspectionAndCryptographicValidation() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureSharedInspectionValidationBudget.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph(new string('x', 4096));
                document.Save();
            }

            string documentDigest = ComputePackagePartSha256Digest(filePath, "/word/document.xml");
            byte[] packageReferenceSignature = CreateSignatureXml(digestValue: documentDigest);
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            string encodedCertificate = Convert.ToBase64String(certificate.Export(X509ContentType.Cert));
            string digest = "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" /><DigestValue>T2ZmaWNlSU1P</DigestValue>";
            byte[] localReferenceSignature = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\">" +
                "<SignedInfo><CanonicalizationMethod Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\" />" +
                "<SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                "<Reference URI=\"#payload\">" + digest + "</Reference></SignedInfo>" +
                "<SignatureValue>AA==</SignatureValue>" +
                "<KeyInfo><X509Data><X509Certificate>" + encodedCertificate + "</X509Certificate></X509Data></KeyInfo>" +
                "<Object Id=\"payload\">" + new string('x', 512) + "</Object></Signature>");
            AddDigitalSignatureMetadata(filePath, packageReferenceSignature, localReferenceSignature);

            long documentPartLength;
            using (var archive = ZipFile.OpenRead(filePath)) {
                documentPartLength = archive.GetEntry("word/document.xml")!.Length;
            }
            using WordDocument loaded = WordDocument.Load(filePath);
            WordSignatureValidationReport bounded = loaded.ValidateSignatures(SecurityProvider, new WordSignatureValidationOptions {
                MaxTotalDigestBytes = checked(documentPartLength + 256L)
            });
            WordSignatureValidationReport allowed = loaded.ValidateSignatures(SecurityProvider, new WordSignatureValidationOptions {
                MaxTotalDigestBytes = checked(documentPartLength + 4096L)
            });

            Assert.Contains(bounded.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
            Assert.DoesNotContain(allowed.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
        }

        [Fact]
        public void Test_DigitalSignature_BoundsPackageDigestWorkAcrossSignatureParts() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignaturePackageDigestBudget.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph(new string('x', 4096));
                document.Save();
            }

            string documentDigest = ComputePackagePartSha256Digest(filePath, "/word/document.xml");
            AddDigitalSignatureMetadata(
                filePath,
                CreateSignatureXml(digestValue: documentDigest),
                signatureCount: 2);

            byte[] packageBytes = File.ReadAllBytes(filePath);
            long documentPartLength;
            using (var archive = new ZipArchive(new MemoryStream(packageBytes), ZipArchiveMode.Read, leaveOpen: false)) {
                documentPartLength = archive.GetEntry("word/document.xml")!.Length;
            }
            using WordprocessingDocument package = WordprocessingDocument.Open(filePath, false);
            OfficePackageSignatureInfo bounded = OfficePackageSignatureInspector.Inspect(
                package,
                package.DigitalSignatureOriginPart,
                hasApplicationSignatureMetadata: true,
                packageBytes,
                maxTotalDigestBytes: documentPartLength);
            OfficePackageSignatureInfo allowed = OfficePackageSignatureInspector.Inspect(
                package,
                package.DigitalSignatureOriginPart,
                hasApplicationSignatureMetadata: true,
                packageBytes,
                maxTotalDigestBytes: checked(documentPartLength * 2L));

            Assert.Equal(2, bounded.SignatureParts.Count);
            Assert.True(bounded.InspectionResourceLimitExceeded);
            Assert.Contains(bounded.UnsupportedDetails, detail =>
                detail.Contains("aggregate digest-work limit", StringComparison.OrdinalIgnoreCase));
            Assert.Single(bounded.SignatureParts, part => part.ParseError == null);
            Assert.Single(bounded.SignatureParts, part => part.ParseError != null &&
                part.ParseError.Contains("aggregate digest-work limit", StringComparison.OrdinalIgnoreCase));
            Assert.All(allowed.SignatureParts, part => {
                Assert.Null(part.ParseError);
                Assert.Single(part.SignedReferences);
            });
        }

        [Fact]
        public void Test_DigitalSignature_BoundsTimestampWorkAcrossSignatureParts() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureTimestampBudget.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Timestamp validation work budget");
                document.Save();
            }
            const string signatureId = "OfficeIMOTimestampBudget";
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\" Id=\"" + signatureId + "\">" +
                "<SignedInfo><CanonicalizationMethod Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\" />" +
                "<SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" /></SignedInfo>" +
                "<SignatureValue>AA==</SignatureValue><Object>" +
                "<xades:QualifyingProperties xmlns:xades=\"http://uri.etsi.org/01903/v1.3.2#\" Target=\"#" + signatureId + "\">" +
                "<xades:UnsignedProperties><xades:UnsignedSignatureProperties><xades:SignatureTimeStamp>" +
                "<xades:EncapsulatedTimeStamp>AA==</xades:EncapsulatedTimeStamp>" +
                "</xades:SignatureTimeStamp></xades:UnsignedSignatureProperties></xades:UnsignedProperties>" +
                "</xades:QualifyingProperties></Object></Signature>");
            AddDigitalSignatureMetadata(filePath, signatureBytes, signatureCount: 2);

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            WordSignatureValidationReport bounded = loaded.ValidateSignatures(SecurityProvider, new WordSignatureValidationOptions {
                ValidateCryptographicSignature = false,
                MaxTimestampTokens = 1
            });
            WordSignatureValidationReport allowed = loaded.ValidateSignatures(SecurityProvider, new WordSignatureValidationOptions {
                ValidateCryptographicSignature = false,
                MaxTimestampTokens = 2
            });

            Assert.Contains(bounded.Diagnostics, finding => finding.Code == "TimestampResourceLimitExceeded");
            Assert.DoesNotContain(allowed.Diagnostics, finding => finding.Code == "TimestampResourceLimitExceeded");
        }

        [Fact]
        public void Test_DigitalSignature_ReportsRelatedCertificateAggregateLimitAsResourceFailure() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureRelatedCertificateBudget.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Related certificate byte budget");
                document.Save();
            }
            AddDigitalSignatureMetadata(filePath, CreateSignatureXml(digestValue: "T2ZmaWNlSU1P"));
            using X509Certificate2 firstCertificate = CreateSelfSignedSigningCertificate("CN=OfficeIMO Related One");
            using X509Certificate2 secondCertificate = CreateSelfSignedSigningCertificate("CN=OfficeIMO Related Two");
            byte[] firstBytes = firstCertificate.Export(X509ContentType.Cert);
            byte[] secondBytes = secondCertificate.Export(X509ContentType.Cert);
            AddRelatedSignatureCertificates(filePath, firstBytes, secondBytes);

            byte[] packageBytes = File.ReadAllBytes(filePath);
            using WordprocessingDocument package = WordprocessingDocument.Open(filePath, false);
            WordSignatureInfo boundedInspection = WordSignatureInspector.Inspect(
                package,
                package.DigitalSignatureOriginPart,
                hasApplicationSignatureMetadata: true,
                packageBytes,
                maxTotalCertificateBytes: firstBytes.LongLength + 1);
            WordSignatureInfo signatureInfo = WordSignatureInspector.Inspect(
                package,
                package.DigitalSignatureOriginPart,
                hasApplicationSignatureMetadata: true,
                packageBytes,
                maxTotalCertificateBytes: firstBytes.LongLength + secondBytes.LongLength);
            IReadOnlyList<WordSignaturePartValidationResult> validation = OfficePackageSignatureValidator.Validate(
                package.DigitalSignatureOriginPart,
                packageBytes,
                signatureInfo,
                SecurityProvider,
                new WordSignatureValidationOptions {
                    ValidateCryptographicSignature = false,
                    MaxTotalCertificateBytes = firstBytes.LongLength + 1
                });

            Assert.Contains(boundedInspection.SignatureParts, part =>
                part.ParseError?.Contains("aggregate certificate limit", StringComparison.OrdinalIgnoreCase) == true);
            Assert.Contains(Assert.Single(validation).Findings, finding =>
                finding.Code == "SignatureResourceLimitExceeded");
            Assert.DoesNotContain(Assert.Single(validation).Findings, finding =>
                finding.Code == "CertificateMalformed");
        }

        [Fact]
        public void Test_DigitalSignature_AcceptsWhitespaceWrappedCertificateAtDecodedByteLimit() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureWhitespaceCertificate.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Whitespace-wrapped signer certificate");
                document.Save();
            }
            using X509Certificate2 signer = CreateSelfSignedSigningCertificate();
            byte[] certificateBytes = signer.Export(X509ContentType.Cert);
            Assert.True(WordDocument.SignPackage(filePath, SecurityProvider, signer).Succeeded);
            string encodedCertificate = Convert.ToBase64String(certificateBytes);
            string wrappedCertificate = Convert.ToBase64String(
                certificateBytes,
                Base64FormattingOptions.InsertLineBreaks);
            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true)) {
                XmlSignaturePart signaturePart = Assert.Single(package.DigitalSignatureOriginPart!.XmlSignatureParts);
                string signatureXml;
                using (var input = new StreamReader(signaturePart.GetStream(FileMode.Open, FileAccess.Read), Encoding.UTF8)) {
                    signatureXml = input.ReadToEnd();
                }
                Assert.Contains(encodedCertificate, signatureXml, StringComparison.Ordinal);
                using var replacement = new MemoryStream(Encoding.UTF8.GetBytes(
                    signatureXml.Replace(encodedCertificate, wrappedCertificate)));
                signaturePart.FeedData(replacement);
            }

            byte[] packageBytes = File.ReadAllBytes(filePath);
            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, false)) {
                WordSignatureInfo inspection = WordSignatureInspector.Inspect(
                    package,
                    package.DigitalSignatureOriginPart,
                    hasApplicationSignatureMetadata: true,
                    packageBytes,
                    maxCertificateBytes: certificateBytes.LongLength,
                    maxTotalCertificateBytes: certificateBytes.LongLength);

                WordSignaturePartInfo part = Assert.Single(inspection.SignatureParts);
                Assert.Null(part.ParseError);
                Assert.Contains(part.X509SubjectNames, subject =>
                    subject.Contains("OfficeIMO Package Signing Test", StringComparison.Ordinal));
            }

            using WordDocument loaded = WordDocument.Load(filePath, new WordLoadOptions {
                AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
            });
            var options = new WordSignatureValidationOptions {
                MaxCertificateBytes = certificateBytes.LongLength,
                MaxTotalCertificateBytes = certificateBytes.LongLength
            };
            options.CertificateValidation.DisableCertificateDownloads = false;
            options.CertificateValidation.ChainEvaluator = static (_, _) => true;
            WordSignatureValidationReport validation = loaded.ValidateSignatures(SecurityProvider, options);

            Assert.DoesNotContain(validation.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
            Assert.Equal(WordSignatureValidationState.Passed, validation.CryptographicStatus);
            Assert.True(validation.IsValidUnderPolicy, string.Join(Environment.NewLine, validation.Findings));
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
            WordSignatureValidationReport validation = loaded.ValidateSignatures(SecurityProvider, new WordSignatureValidationOptions {
                MaxTotalDigestBytes = 1024
            });

            Assert.Contains(validation.Diagnostics, finding => finding.Code == "SignatureResourceLimitExceeded");
        }

        [Fact]
        public void Test_DigitalSignature_BoundsPendingDomSerializationBeforeSnapshotAllocation() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignaturePendingDomLimit.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed content");
                document.Save();
            }
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            Assert.True(WordDocument.SignPackage(filePath, SecurityProvider, certificate).Succeeded);

            using WordDocument loaded = WordDocument.Load(filePath);
            loaded.AddParagraph(new string('x', 16_384));
            WordSignatureValidationReport validation = loaded.ValidateSignatures(SecurityProvider,
                new WordSignatureValidationOptions {
                    MaxPartBytes = 4096,
                    MaxPackageBytes = 16L * 1024 * 1024
                });

            Assert.Contains(validation.Diagnostics, finding =>
                finding.Code == "SignatureResourceLimitExceeded" &&
                finding.Message.Contains("pending package part", StringComparison.OrdinalIgnoreCase));
        }

#if NETFRAMEWORK
        [Fact]
        public void Test_DigitalSignature_LegacyRuntimeBoundsLivePackageSerialization() {
            string filePath = Path.Combine(_directoryWithFiles, "WordDigitalSignatureLegacyLivePackageLimit.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed content");
                document.Save();
            }
            using X509Certificate2 certificate = CreateSelfSignedSigningCertificate();
            Assert.True(WordDocument.SignPackage(filePath, SecurityProvider, certificate).Succeeded);
            long originalPackageBytes = new FileInfo(filePath).Length;

            using WordDocument loaded = WordDocument.Load(filePath);
            byte[] randomBytes = new byte[256 * 1024];
            new Random(17).NextBytes(randomBytes);
            loaded.AddParagraph(Convert.ToBase64String(randomBytes));
            WordSignatureValidationReport validation = loaded.ValidateSignatures(
                SecurityProvider,
                new WordSignatureValidationOptions {
                    MaxPartBytes = 2L * 1024 * 1024,
                    MaxPackageBytes = originalPackageBytes + 1024L,
                    MaxTotalDigestBytes = 2L * 1024 * 1024
                });

            Assert.Contains(validation.Diagnostics, finding =>
                finding.Code == "SignatureResourceLimitExceeded" &&
                finding.Message.Contains("validation-snapshot limit", StringComparison.OrdinalIgnoreCase));
        }
#endif

        private static void AddRelatedSignatureCertificates(string filePath, params byte[][] certificates) {
            using WordprocessingDocument package = WordprocessingDocument.Open(filePath, true);
            XmlSignaturePart signaturePart = package.DigitalSignatureOriginPart!.XmlSignatureParts.Single();
            foreach (byte[] certificate in certificates) {
                ExtendedPart certificatePart = signaturePart.AddExtendedPart(
                    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate",
                    "application/vnd.openxmlformats-package.digital-signature-certificate",
                    "cer");
                using var stream = new MemoryStream(certificate);
                certificatePart.FeedData(stream);
            }
        }
    }
}

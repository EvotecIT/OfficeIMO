using System.IO.Packaging;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Security;

namespace OfficeIMO.Word {
    internal static class WordMacroProjectSignatureInspector {
        internal sealed class InspectionBudget : IDisposable {
            internal InspectionBudget(
                WordMacroProjectSignatureInspectionOptions options,
                IOfficeSecurityProvider? securityProvider) {
                CmsSession = options.ValidateCms && securityProvider != null
                    ? securityProvider.CreateCmsVerificationSession(options.CmsVerification)
                    : null;
            }

            internal ICmsVerificationSession? CmsSession { get; }

            public void Dispose() => CmsSession?.Dispose();
        }

        internal const string LegacyRelationship = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";
        internal const string AgileRelationship = "http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile";
        internal const string AgileRelationshipCompatibility = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignatureAgile";
        internal const string V3Relationship = "http://schemas.microsoft.com/office/2020/07/relationships/vbaProjectSignatureV3";

        private const string LegacyContentType = "application/vnd.ms-office.vbaProjectSignature";
        private const string AgileContentType = "application/vnd.ms-office.vbaProjectSignatureAgile";
        private const string V3ContentType = "application/vnd.ms-office.vbaProjectSignatureV3";

        internal static WordMacroProjectSignatureInfo Inspect(
            string filePath,
            WordMacroProjectSignatureInspectionOptions? options = null) {
            options ??= new WordMacroProjectSignatureInspectionOptions();
            using var operationBudget = new InspectionBudget(options, securityProvider: null);
            return Inspect(filePath, options, operationBudget, profileFilter: null, validateCmsOverride: false);
        }

        internal static WordMacroProjectSignatureInfo Inspect(
            string filePath,
            IOfficeSecurityProvider securityProvider,
            WordMacroProjectSignatureInspectionOptions? options = null) {
            if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
            options ??= new WordMacroProjectSignatureInspectionOptions();
            using var operationBudget = new InspectionBudget(options, securityProvider);
            return Inspect(filePath, options, operationBudget);
        }

        internal static WordMacroProjectSignatureInfo Inspect(
            string filePath,
            WordMacroProjectSignatureInspectionOptions? options,
            InspectionBudget? operationBudget,
            WordMacroProjectSignatureProfile? profileFilter = null,
            bool? validateCmsOverride = null,
            Action? afterSnapshotCaptured = null) {
            options ??= new WordMacroProjectSignatureInspectionOptions();
            ValidateOptions(options);
            operationBudget ??= new InspectionBudget(options, securityProvider: null);

            string fullPath = NormalizePath(filePath);
            bool macroEnabled = IsMacroEnabledPath(fullPath);
            var findings = new List<WordMacroProjectSignatureFinding>();
            if (!macroEnabled) {
                findings.Add(Finding("MacroEnabledFormatRequired", WordSignatureValidationState.Failed,
                    "VBA signature inspection supports saved DOCM and DOTM packages only."));
            }
            if (!File.Exists(fullPath)) {
                findings.Add(Finding("FileNotFound", WordSignatureValidationState.Failed,
                    "The Word package does not exist: " + fullPath));
                return Empty(fullPath, macroEnabled, findings);
            }

            var file = new FileInfo(fullPath);
            if (file.Length > options.PackageSecurity.MaxPackageBytes) {
                findings.Add(Finding("PackageByteLimitExceeded", WordSignatureValidationState.Failed,
                    "The Word package exceeds the configured " + options.PackageSecurity.MaxPackageBytes + " byte limit."));
                return Empty(fullPath, macroEnabled, findings);
            }

            Uri? discoveredVbaUri = null;
            long? discoveredVbaLength = null;
            string? discoveredVbaHash = null;
            try {
                byte[] packageBytes;
                using (FileStream packageStream = File.OpenRead(fullPath)) {
                    packageBytes = OfficePackageSecurityInspector.ReadAndValidate(
                        packageStream,
                        options.PackageSecurity);
                }
                afterSnapshotCaptured?.Invoke();
                using (var documentStream = new MemoryStream(packageBytes, writable: false))
                using (WordprocessingDocument document = WordprocessingDocument.Open(documentStream, false)) {
                    discoveredVbaUri = document.MainDocumentPart?.VbaProjectPart?.Uri;
                }
                if (discoveredVbaUri == null) {
                    findings.Add(Finding("MacroProjectNotPresent", WordSignatureValidationState.NotPresent,
                        "The Word package does not contain a VBA project part."));
                    return Empty(fullPath, macroEnabled, findings);
                }

                using (var packageStream = new MemoryStream(packageBytes, writable: false))
                using (Package package = Package.Open(packageStream, FileMode.Open, FileAccess.Read)) {
                    if (!package.PartExists(discoveredVbaUri)) {
                        findings.Add(Finding("MacroProjectPartMissing", WordSignatureValidationState.Failed,
                            "The VBA project relationship target is missing from the package."));
                        return new WordMacroProjectSignatureInfo(fullPath, macroEnabled, true,
                            discoveredVbaUri.ToString(), null, null,
                            Array.Empty<WordMacroProjectSignaturePartInfo>(), findings);
                    }

                    PackagePart vbaPart = package.GetPart(discoveredVbaUri);
                    discoveredVbaHash = HashPart(vbaPart, options.MaxMacroProjectBytes, out long vbaLength);
                    discoveredVbaLength = vbaLength;
                    IReadOnlyList<WordMacroProjectSignaturePartInfo> signatures = InspectSignatureRelationships(
                        package, vbaPart, options, operationBudget, profileFilter, validateCmsOverride, findings);
                    return new WordMacroProjectSignatureInfo(fullPath, macroEnabled, true,
                        discoveredVbaUri.ToString(), discoveredVbaLength, discoveredVbaHash, signatures, findings);
                }
            } catch (OfficePackageSecurityException exception) {
                findings.Add(Finding("PackageSecurity" + exception.Rule, WordSignatureValidationState.Failed,
                    "The Word package failed the shared security policy before Open XML parsing. " + exception.Message));
                return Empty(fullPath, macroEnabled, findings);
            } catch (Exception exception) when (IsInspectionException(exception)) {
                findings.Add(Finding("MacroSignatureInspectionFailed", WordSignatureValidationState.Failed,
                    "The VBA signature package structure could not be inspected. " + exception.Message));
                return PartialOrEmpty(
                    fullPath,
                    macroEnabled,
                    discoveredVbaUri,
                    discoveredVbaLength,
                    discoveredVbaHash,
                    findings);
            }
        }

        internal static void ValidateOptions(WordMacroProjectSignatureInspectionOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (options.MaxMacroProjectBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxMacroProjectBytes));
            if (options.MaxSignatureBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxSignatureBytes));
            if (options.MaxTotalSignatureBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTotalSignatureBytes));
            if (options.MaxRelationships <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxRelationships));
            OfficePackageSecurityInspector.ValidateOptions(options.PackageSecurity);
            ValidateCmsOptions(options.CmsVerification);
        }

        internal static bool IsMacroEnabledPath(string filePath) {
            string extension = Path.GetExtension(filePath);
            return string.Equals(extension, ".docm", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".dotm", StringComparison.OrdinalIgnoreCase);
        }

        private static IReadOnlyList<WordMacroProjectSignaturePartInfo> InspectSignatureRelationships(
            Package package,
            PackagePart vbaPart,
            WordMacroProjectSignatureInspectionOptions options,
            InspectionBudget operationBudget,
            WordMacroProjectSignatureProfile? profileFilter,
            bool? validateCmsOverride,
            ICollection<WordMacroProjectSignatureFinding> findings) {
            var signatures = new List<WordMacroProjectSignaturePartInfo>();
            var profiles = new HashSet<WordMacroProjectSignatureProfile>();
            long aggregateBytes = 0;
            int relationshipCount = 0;
            foreach (PackageRelationship relationship in vbaPart.GetRelationships()) {
                relationshipCount++;
                if (relationshipCount > options.MaxRelationships) {
                    throw new InvalidDataException("The VBA project has more than " + options.MaxRelationships + " relationships.");
                }
                if (!TryGetProfile(relationship.RelationshipType, out WordMacroProjectSignatureProfile profile)) {
                    continue;
                }
                if (profileFilter.HasValue && profile != profileFilter.Value) continue;
                if (relationship.TargetMode != TargetMode.Internal) {
                    findings.Add(Finding("MacroSignatureExternalTarget", WordSignatureValidationState.Failed,
                        "A VBA signature relationship targets an external resource.", profile));
                    continue;
                }

                Uri partUri = PackUriHelper.ResolvePartUri(vbaPart.Uri, relationship.TargetUri);
                if (!package.PartExists(partUri)) {
                    findings.Add(Finding("MacroSignaturePartMissing", WordSignatureValidationState.Failed,
                        "The " + profile + " VBA signature relationship target is missing.", profile));
                    continue;
                }
                if (!profiles.Add(profile)) {
                    findings.Add(Finding("DuplicateMacroSignatureProfile", WordSignatureValidationState.Failed,
                        "More than one " + profile + " VBA signature relationship is present.", profile));
                    continue;
                }

                PackagePart signaturePart = package.GetPart(partUri);
                byte[] encodedPart = ReadPart(signaturePart, options.MaxSignatureBytes);
                long length = encodedPart.LongLength;
                aggregateBytes = checked(aggregateBytes + length);
                if (aggregateBytes > options.MaxTotalSignatureBytes) {
                    throw new InvalidDataException("The aggregate VBA signature bytes exceed the configured " +
                        options.MaxTotalSignatureBytes + " byte limit.");
                }

                signatures.Add(InspectSignaturePart(signaturePart, relationship.RelationshipType,
                    profile, encodedPart, options, operationBudget, validateCmsOverride));
            }

            signatures.Sort((left, right) => left.Profile.CompareTo(right.Profile));
            return signatures;
        }

        private static WordMacroProjectSignaturePartInfo InspectSignaturePart(
            PackagePart part,
            string relationshipType,
            WordMacroProjectSignatureProfile profile,
            byte[] encodedPart,
            WordMacroProjectSignatureInspectionOptions options,
            InspectionBudget operationBudget,
            bool? validateCmsOverride) {
            long length = encodedPart.LongLength;
            var findings = new List<WordMacroProjectSignatureFinding>();
            string expectedContentType = GetExpectedContentType(profile);
            if (!string.Equals(part.ContentType, expectedContentType, StringComparison.OrdinalIgnoreCase)) {
                findings.Add(Finding("MacroSignatureContentTypeUnexpected", WordSignatureValidationState.Failed,
                    "The " + profile + " signature part content type is '" + part.ContentType +
                    "' instead of '" + expectedContentType + "'.", profile));
            }
            if (!(validateCmsOverride ?? options.ValidateCms)) {
                return CreatePartInfo(profile, part, relationshipType, length, false,
                    WordSignatureValidationState.NotChecked, WordSignatureValidationState.NotChecked,
                    WordSignatureValidationState.NotChecked, WordSignatureValidationState.NotChecked,
                    null, null, findings);
            }

            if (operationBudget.CmsSession == null) {
                findings.Add(Finding("SecurityProviderRequired", WordSignatureValidationState.NotChecked,
                    "CMS signature, certificate, revocation, and timestamp validation requires an explicit OfficeIMO.Security provider.",
                    profile));
                return CreatePartInfo(profile, part, relationshipType, length, false,
                    WordSignatureValidationState.NotChecked, WordSignatureValidationState.NotChecked,
                    WordSignatureValidationState.NotChecked, WordSignatureValidationState.NotChecked,
                    null, null, findings);
            }

            if (!TryExtractCms(encodedPart, options.CmsVerification.MaxEncodedBytes,
                    out byte[] cmsBytes, out string parseDetail)) {
                findings.Add(Finding("MacroSignatureContainerMalformed", WordSignatureValidationState.Failed,
                    parseDetail, profile));
                return CreatePartInfo(profile, part, relationshipType, length, false,
                    WordSignatureValidationState.Failed, WordSignatureValidationState.NotChecked,
                    WordSignatureValidationState.NotChecked, WordSignatureValidationState.NotChecked,
                    null, null, findings);
            }

            CmsVerificationResult cms = operationBudget.CmsSession.Verify(
                cmsBytes,
                CertificateValidationPurpose.DocumentSigning);
            bool hasSingleSigner = cms.Signers.Count == 1;
            if (!hasSingleSigner) {
                findings.Add(Finding(
                    "MacroSignatureSignerCountInvalid",
                    WordSignatureValidationState.Failed,
                    "A VBA signature profile must contain exactly one CMS signer; the container declares " +
                    cms.Signers.Count + ".",
                    profile));
            }
            CmsSignerVerificationResult? signer = hasSingleSigner ? cms.Signers[0] : null;
            foreach (SecurityFinding finding in cms.Findings.Concat(cms.Signers.SelectMany(item => item.Findings))) {
                findings.Add(Finding(finding.Code, Map(finding.Severity), finding.Message, profile));
            }
            WordSignatureValidationState crypto = hasSingleSigner && cms.IsCryptographicallyValid
                ? WordSignatureValidationState.Passed
                : WordSignatureValidationState.Failed;
            WordSignatureValidationState chain = signer == null
                ? WordSignatureValidationState.NotPresent
                : Map(signer.CertificateValidation.ChainStatus);
            WordSignatureValidationState revocation = signer == null
                ? WordSignatureValidationState.NotPresent
                : Map(signer.CertificateValidation.RevocationStatus);
            WordSignatureValidationState timestamp = signer == null
                ? WordSignatureValidationState.NotPresent
                : Map(signer.TimestampStatus);
            return CreatePartInfo(profile, part, relationshipType, length, cms.Parsed,
                crypto, chain, revocation, timestamp, signer, cms.AuthenticodeIndirectData, findings);
        }

        private static WordMacroProjectSignaturePartInfo CreatePartInfo(
            WordMacroProjectSignatureProfile profile,
            PackagePart part,
            string relationshipType,
            long length,
            bool cmsParsed,
            WordSignatureValidationState cryptographicStatus,
            WordSignatureValidationState chainStatus,
            WordSignatureValidationState revocationStatus,
            WordSignatureValidationState timestampStatus,
            CmsSignerVerificationResult? signer,
            AuthenticodeIndirectDataInfo? authenticode,
            IReadOnlyList<WordMacroProjectSignatureFinding> findings) {
            return new WordMacroProjectSignaturePartInfo(profile, part.Uri.ToString(), relationshipType,
                part.ContentType, length, cmsParsed, cryptographicStatus, chainStatus, revocationStatus,
                timestampStatus, signer?.Subject, signer?.Thumbprint, signer?.DigestAlgorithmOid,
                signer?.SignatureAlgorithmOid, authenticode?.DigestAlgorithmOid,
                authenticode?.Digest == null ? null : (byte[])authenticode.Digest.Clone(),
                signer?.SigningTime, signer?.TimestampTime, findings);
        }

        private static bool TryExtractCms(byte[] encodedPart, long maxCmsBytes, out byte[] cms, out string detail) {
            const int digSigBlobHeaderLength = 8;
            const int digSigInfoSerializedHeaderLength = 36;
            cms = Array.Empty<byte>();
            detail = string.Empty;
            if (encodedPart.Length < digSigInfoSerializedHeaderLength) {
                detail = "The VBA signature part is shorter than DigSigInfoSerialized.";
                return false;
            }
            uint signatureLength = ReadUInt32(encodedPart, 0);
            uint serializedSignatureOffset = ReadUInt32(encodedPart, 4);
            if (serializedSignatureOffset < digSigBlobHeaderLength + digSigInfoSerializedHeaderLength) {
                detail = "The VBA signature CMS offset or length is outside the bounded signature part.";
                return false;
            }
            uint partSignatureOffset = serializedSignatureOffset - digSigBlobHeaderLength;
            if (signatureLength == 0 || signatureLength > maxCmsBytes ||
                partSignatureOffset > encodedPart.Length ||
                signatureLength > encodedPart.Length - partSignatureOffset) {
                detail = "The VBA signature CMS offset or length is outside the bounded signature part.";
                return false;
            }
            cms = new byte[signatureLength];
            Buffer.BlockCopy(encodedPart, checked((int)partSignatureOffset), cms, 0, checked((int)signatureLength));
            return true;
        }

        private static void ValidateCmsOptions(CmsVerificationOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (options.MaxEncodedBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxEncodedBytes));
            if (options.MaxContentBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxContentBytes));
            if (options.MaxSigners <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxSigners));
            if (options.MaxCertificates <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxCertificates));
            if (options.MaxTimestampTokens <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTimestampTokens));
            if (options.MaxTimestampTokenBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTimestampTokenBytes));
            if (options.MaxTotalTimestampBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTotalTimestampBytes));
        }

        private static uint ReadUInt32(byte[] bytes, int offset) =>
            (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24);

        private static byte[] ReadPart(PackagePart part, long maximum) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            if (stream.CanSeek && stream.Length > maximum) throw new InvalidDataException("The package part exceeds its byte limit.");
            using var output = new MemoryStream();
            CopyBounded(stream, output, maximum);
            return output.ToArray();
        }

        private static string HashPart(PackagePart part, long maximum, out long length) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var bounded = new BoundedHashStream(maximum);
            CopyBounded(stream, bounded, maximum);
            length = bounded.Length;
            return bounded.GetHash();
        }

        private static void CopyBounded(Stream input, Stream output, long maximum) {
            byte[] buffer = new byte[81920];
            long total = 0;
            while (true) {
                int read = input.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                total = checked(total + read);
                if (total > maximum) throw new InvalidDataException("The package part exceeds its configured byte limit.");
                output.Write(buffer, 0, read);
            }
        }

        private static bool TryGetProfile(string relationshipType, out WordMacroProjectSignatureProfile profile) {
            if (string.Equals(relationshipType, LegacyRelationship, StringComparison.OrdinalIgnoreCase)) {
                profile = WordMacroProjectSignatureProfile.Legacy;
                return true;
            }
            if (string.Equals(relationshipType, AgileRelationship, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(relationshipType, AgileRelationshipCompatibility, StringComparison.OrdinalIgnoreCase)) {
                profile = WordMacroProjectSignatureProfile.Agile;
                return true;
            }
            if (string.Equals(relationshipType, V3Relationship, StringComparison.OrdinalIgnoreCase)) {
                profile = WordMacroProjectSignatureProfile.V3;
                return true;
            }
            profile = WordMacroProjectSignatureProfile.Unknown;
            return false;
        }

        private static string GetExpectedContentType(WordMacroProjectSignatureProfile profile) {
            switch (profile) {
                case WordMacroProjectSignatureProfile.Legacy: return LegacyContentType;
                case WordMacroProjectSignatureProfile.Agile: return AgileContentType;
                case WordMacroProjectSignatureProfile.V3: return V3ContentType;
                default: return "application/octet-stream";
            }
        }

        private static WordSignatureValidationState Map(SecurityValidationStatus status) {
            switch (status) {
                case SecurityValidationStatus.Valid: return WordSignatureValidationState.Passed;
                case SecurityValidationStatus.Invalid: return WordSignatureValidationState.Failed;
                case SecurityValidationStatus.NotPerformed: return WordSignatureValidationState.NotChecked;
                default: return WordSignatureValidationState.Unsupported;
            }
        }

        private static WordSignatureValidationState Map(SecurityFindingSeverity severity) =>
            severity == SecurityFindingSeverity.Error
                ? WordSignatureValidationState.Failed
                : WordSignatureValidationState.NotChecked;

        private static WordMacroProjectSignatureFinding Finding(
            string code,
            WordSignatureValidationState state,
            string message,
            WordMacroProjectSignatureProfile? profile = null) =>
            new WordMacroProjectSignatureFinding(code, state, message, profile);

        private static WordMacroProjectSignatureInfo Empty(
            string filePath,
            bool macroEnabled,
            IReadOnlyList<WordMacroProjectSignatureFinding> findings) =>
            new WordMacroProjectSignatureInfo(filePath, macroEnabled, false, null, null, null,
                Array.Empty<WordMacroProjectSignaturePartInfo>(), findings);

        private static WordMacroProjectSignatureInfo PartialOrEmpty(
            string filePath,
            bool macroEnabled,
            Uri? vbaUri,
            long? vbaLength,
            string? vbaHash,
            IReadOnlyList<WordMacroProjectSignatureFinding> findings) =>
            vbaUri == null
                ? Empty(filePath, macroEnabled, findings)
                : new WordMacroProjectSignatureInfo(filePath, macroEnabled, true, vbaUri.ToString(),
                    vbaLength, vbaHash, Array.Empty<WordMacroProjectSignaturePartInfo>(), findings);

        private static string NormalizePath(string filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A Word package path is required.", nameof(filePath));
            return Path.GetFullPath(filePath);
        }

        private static bool IsInspectionException(Exception exception) =>
            exception is IOException || exception is InvalidDataException || exception is UnauthorizedAccessException ||
            exception is NotSupportedException || exception is ArgumentException || exception is OverflowException ||
            exception is System.Xml.XmlException || exception is System.Security.Cryptography.CryptographicException;

        private sealed class BoundedHashStream : Stream {
            private readonly long _maximum;
            private readonly HashAlgorithm _hash = SHA256.Create();
            private long _length;
            private bool _finished;

            internal BoundedHashStream(long maximum) => _maximum = maximum;

            public override void Write(byte[] buffer, int offset, int count) {
                if (_finished) throw new InvalidOperationException("The hash is already finalized.");
                _length = checked(_length + count);
                if (_length > _maximum) throw new InvalidDataException("The VBA project exceeds its configured byte limit.");
                _hash.TransformBlock(buffer, offset, count, null, 0);
            }

            internal string GetHash() {
                if (!_finished) {
                    _hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                    _finished = true;
                }
                return BitConverter.ToString(_hash.Hash ?? Array.Empty<byte>()).Replace("-", string.Empty);
            }

            protected override void Dispose(bool disposing) {
                if (disposing) _hash.Dispose();
                base.Dispose(disposing);
            }

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => _length;
            public override long Position { get => _length; set => throw new NotSupportedException(); }
            public override void Flush() { }
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
        }
    }
}

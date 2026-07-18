using OfficeIMO.Drawing.Internal;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeIMO.Drawing {
    /// <summary>Inspects and validates Office package structure without executing active content.</summary>
    public static partial class OfficePackageSecurityInspector {
        private const string RelationshipSuffix = ".rels";

        /// <summary>Inspects package bytes using secure structural defaults.</summary>
        public static OfficePackageSecurityReport Inspect(byte[] packageBytes,
            OfficePackageSecurityOptions? options = null) {
            if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
            OfficePackageSecurityOptions resolved = options ?? OfficePackageSecurityOptions.SecureDefaults;
            ValidateOptions(resolved);

            var findings = new List<OfficePackageSecurityFinding>();
            AddLimitFinding(findings, OfficePackageSecurityRule.PackageSize, packageBytes.Length,
                resolved.MaxPackageBytes, "Package size");

            if (packageBytes.LongLength > resolved.MaxPackageBytes) {
                OfficePackageContainerKind oversizedKind = HasZipSignature(packageBytes)
                    ? OfficePackageContainerKind.OpenXml
                    : OfficeCompoundDocumentDetector.HasCompoundSignature(packageBytes)
                        ? OfficePackageContainerKind.CompoundBinary
                        : OfficePackageContainerKind.Unknown;
                return new OfficePackageSecurityReport(packageBytes.LongLength, oversizedKind,
                    0, 0, 0, 0, 0, 0, 0, 0, 0, findings.ToArray());
            }

            if (HasZipSignature(packageBytes)) {
                return InspectZip(packageBytes, resolved, findings);
            }

            if (OfficeCompoundDocumentDetector.HasCompoundSignature(packageBytes)) {
                return InspectCompound(packageBytes, resolved, findings);
            }

            return new OfficePackageSecurityReport(packageBytes.Length, OfficePackageContainerKind.Unknown,
                0, 0, 0, 0, 0, 0, 0, 0, 0, findings.ToArray());
        }

        /// <summary>Reads and inspects a caller-owned stream while preserving its position when seekable.</summary>
        public static OfficePackageSecurityReport Inspect(Stream source,
            OfficePackageSecurityOptions? options = null) {
            OfficePackageSecurityOptions resolved = options ?? OfficePackageSecurityOptions.SecureDefaults;
            byte[] bytes = ReadSource(source, resolved);
            return Inspect(bytes, resolved);
        }

        /// <summary>Validates package bytes and throws a typed exception for the first rejected finding.</summary>
        public static OfficePackageSecurityReport Validate(byte[] packageBytes,
            OfficePackageSecurityOptions? options = null) {
            OfficePackageSecurityReport report = Inspect(packageBytes, options);
            ThrowFirstError(report);
            return report;
        }

        /// <summary>Reads and validates a caller-owned stream while preserving its position when seekable.</summary>
        public static OfficePackageSecurityReport Validate(Stream source,
            OfficePackageSecurityOptions? options = null) {
            OfficePackageSecurityOptions resolved = options ?? OfficePackageSecurityOptions.SecureDefaults;
            byte[] bytes = ReadSource(source, resolved);
            return Validate(bytes, resolved);
        }

        internal static byte[] ReadAndValidate(Stream source, OfficePackageSecurityOptions options) {
            byte[] bytes = ReadSource(source, options);
            Validate(bytes, options);
            return bytes;
        }

        internal static byte[] ReadBounded(Stream source, OfficePackageSecurityOptions options) =>
            ReadSource(source, options);

        internal static async Task<byte[]> ReadAndValidateAsync(Stream source,
            OfficePackageSecurityOptions options, CancellationToken cancellationToken) {
            ValidateOptions(options);
            try {
                byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(source, cancellationToken,
                    options.MaxPackageBytes).ConfigureAwait(false);
                Validate(bytes, options);
                return bytes;
            } catch (InvalidDataException exception) {
                throw CreateSourceSizeException(source, options, exception);
            }
        }

        internal static async Task<byte[]> ReadBoundedAsync(Stream source,
            OfficePackageSecurityOptions options, CancellationToken cancellationToken) {
            ValidateOptions(options);
            try {
                return await OfficeStreamReader.ReadAllBytesAsync(source, cancellationToken,
                    options.MaxPackageBytes).ConfigureAwait(false);
            } catch (InvalidDataException exception) {
                throw CreateSourceSizeException(source, options, exception);
            }
        }

        internal static void ValidateSourceSize(long sourceBytes, OfficePackageSecurityOptions options) {
            ValidateOptions(options);
            if (sourceBytes > options.MaxPackageBytes) {
                throw new OfficePackageSecurityException(OfficePackageSecurityRule.PackageSize,
                    $"Package size {sourceBytes} bytes exceeds the configured maximum of {options.MaxPackageBytes} bytes.",
                    sourceBytes, options.MaxPackageBytes);
            }
        }

        private static byte[] ReadSource(Stream source, OfficePackageSecurityOptions options) {
            ValidateOptions(options);
            try {
                return OfficeStreamReader.ReadAllBytes(source, options.MaxPackageBytes);
            } catch (InvalidDataException exception) {
                throw CreateSourceSizeException(source, options, exception);
            }
        }

        private static OfficePackageSecurityException CreateSourceSizeException(Stream source,
            OfficePackageSecurityOptions options, InvalidDataException inner) {
            double observed = source != null && source.CanSeek ? source.Length : options.MaxPackageBytes + 1D;
            return new OfficePackageSecurityException(OfficePackageSecurityRule.PackageSize,
                $"Package source exceeds the configured maximum of {options.MaxPackageBytes} bytes. {inner.Message}",
                observed, options.MaxPackageBytes);
        }

        private static OfficePackageSecurityReport InspectZip(byte[] bytes,
            OfficePackageSecurityOptions options, List<OfficePackageSecurityFinding> findings) {
            int partCount = 0;
            long totalBytes = 0;
            long largestPart = 0;
            double highestRatio = 0;
            int externalCount = 0;
            int signatureCount = 0;
            var partNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var macroParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var embeddedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var activeXParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var relationshipParts = new List<ZipXmlPart>();
            ZipXmlPart? contentTypesPart = null;

            OfficeArchiveSafety.ZipCentralDirectoryScanResult centralDirectory =
                OfficeArchiveSafety.ScanZipCentralDirectory(bytes,
                    options.MaxPartCount);
            if (!centralDirectory.IsValid) {
                findings.Add(Error(OfficePackageSecurityRule.MalformedPackage,
                    centralDirectory.Error ?? "The ZIP central directory is malformed."));
                return new OfficePackageSecurityReport(bytes.Length, OfficePackageContainerKind.OpenXml,
                    0, 0, 0, 0, 0, 0, 0, 0, 0, findings.ToArray());
            }
            if (centralDirectory.LimitExceeded) {
                findings.Add(Error(OfficePackageSecurityRule.PartCount,
                    $"ZIP entry count exceeds the configured maximum of {options.MaxPartCount} before package parts are opened.",
                    observedValue: centralDirectory.EntryCount, limit: options.MaxPartCount));
                return new OfficePackageSecurityReport(bytes.Length, OfficePackageContainerKind.OpenXml,
                    0, 0, 0, 0, 0, 0, 0, 0, 0, findings.ToArray());
            }

            try {
                using var source = new MemoryStream(bytes, writable: false);
                using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: false);
                foreach (ZipArchiveEntry entry in archive.Entries) {
                    if (IsDirectory(entry)) continue;
                    partCount++;
                    string partName = NormalizePartName(entry.FullName);
                    long partBytes = entry.Length;
                    long compressedBytes = entry.CompressedLength;
                    totalBytes = checked(totalBytes + partBytes);
                    if (partBytes > largestPart) largestPart = partBytes;

                    double ratio = partBytes == 0
                        ? 0
                        : compressedBytes == 0 ? double.PositiveInfinity : (double)partBytes / compressedBytes;
                    if (ratio > highestRatio) highestRatio = ratio;

                    if (!partNames.Add(partName)) {
                        findings.Add(Error(OfficePackageSecurityRule.DuplicatePartName,
                            $"Package contains a duplicate or case-ambiguous part name '{partName}'.", partName));
                    }
                    bool unsafePartName = IsUnsafePartName(entry.FullName);
                    if (unsafePartName) {
                        findings.Add(Error(OfficePackageSecurityRule.UnsafePartName,
                            $"Package part name '{partName}' is unsafe.", partName));
                    }
                    if (partBytes > options.MaxPartUncompressedBytes) {
                        findings.Add(Error(OfficePackageSecurityRule.PartSize,
                            $"Part '{partName}' is {partBytes} bytes and exceeds the configured maximum of {options.MaxPartUncompressedBytes} bytes.",
                            partName, partBytes, options.MaxPartUncompressedBytes));
                    }
                    if (ratio > options.MaxCompressionRatio) {
                        findings.Add(Error(OfficePackageSecurityRule.CompressionRatio,
                            $"Part '{partName}' has compression ratio {FormatRatio(ratio)} and exceeds the configured maximum of {FormatRatio(options.MaxCompressionRatio)}.",
                            partName, ratio, options.MaxCompressionRatio));
                    }

                    string lowerName = partName.ToLowerInvariant();
                    if (IsMacroPart(lowerName)) macroParts.Add(partName);
                    if (IsEmbeddedPart(lowerName)) embeddedParts.Add(partName);
                    if (IsActiveXPart(lowerName)) activeXParts.Add(partName);
                    if (IsDigitalSignaturePart(lowerName)) signatureCount++;

                    bool safeToInspect = !unsafePartName
                        && partBytes <= options.MaxPartUncompressedBytes
                        && ratio <= options.MaxCompressionRatio;
                    if (safeToInspect) {
                        if (string.Equals(partName, "/[Content_Types].xml", StringComparison.OrdinalIgnoreCase)) {
                            contentTypesPart = new ZipXmlPart(entry, partName);
                        } else if (entry.FullName.EndsWith(RelationshipSuffix, StringComparison.OrdinalIgnoreCase)) {
                            relationshipParts.Add(new ZipXmlPart(entry, partName));
                        }
                    }
                }

                if (partCount <= options.MaxPartCount && totalBytes <= options.MaxTotalUncompressedBytes) {
                    if (contentTypesPart != null) {
                        InspectContentTypes(contentTypesPart, partNames, macroParts, embeddedParts, activeXParts, findings);
                    }
                    foreach (ZipXmlPart relationshipPart in relationshipParts) {
                        externalCount += InspectRelationships(
                            relationshipPart.Entry,
                            relationshipPart.PartName,
                            macroParts,
                            embeddedParts,
                            activeXParts,
                            findings);
                    }
                }
            } catch (Exception exception) when (exception is InvalidDataException || exception is IOException
                || exception is NotSupportedException || exception is OverflowException) {
                findings.Add(Error(OfficePackageSecurityRule.MalformedPackage,
                    $"The ZIP package could not be inspected safely. {exception.Message}"));
            }

            AddLimitFinding(findings, OfficePackageSecurityRule.PartCount, partCount,
                options.MaxPartCount, "Package part count");
            AddLimitFinding(findings, OfficePackageSecurityRule.TotalUncompressedSize, totalBytes,
                options.MaxTotalUncompressedBytes, "Total uncompressed package size");
            int macroCount = macroParts.Count;
            int embeddedCount = embeddedParts.Count;
            int activeXCount = activeXParts.Count;
            AddPolicyFinding(findings, options.Macros, OfficePackageSecurityRule.Macros, macroCount,
                "VBA project part");
            AddPolicyFinding(findings, options.EmbeddedPayloads, OfficePackageSecurityRule.EmbeddedPayloads,
                embeddedCount, "embedded payload part");
            AddPolicyFinding(findings, options.ActiveX, OfficePackageSecurityRule.ActiveX, activeXCount,
                "ActiveX part");
            AddPolicyFinding(findings, options.ExternalRelationships,
                OfficePackageSecurityRule.ExternalRelationships, externalCount, "external relationship");

            return new OfficePackageSecurityReport(bytes.Length, OfficePackageContainerKind.OpenXml,
                partCount, totalBytes, largestPart, highestRatio, macroCount, embeddedCount, activeXCount,
                externalCount, signatureCount, findings.ToArray());
        }

        private static OfficePackageSecurityReport InspectCompound(byte[] bytes,
            OfficePackageSecurityOptions options, List<OfficePackageSecurityFinding> findings) {
            int inspectionEntryLimit = options.MaxPartCount >= int.MaxValue - 4
                ? int.MaxValue
                : options.MaxPartCount + 4;
            using var source = new MemoryStream(bytes, writable: false);
            if (!OfficeCompoundFileReader.TryInspectDirectory(source, options.MaxPackageBytes,
                inspectionEntryLimit, out IReadOnlyList<OfficeCompoundFileEntry> entries, out string? error)) {
                OfficePackageSecurityRule rule = error != null
                    && error.IndexOf("entry count exceeds", StringComparison.OrdinalIgnoreCase) >= 0
                    ? OfficePackageSecurityRule.PartCount
                    : OfficePackageSecurityRule.MalformedPackage;
                double? observed = rule == OfficePackageSecurityRule.PartCount
                    ? options.MaxPartCount + 1D
                    : (double?)null;
                double? limit = rule == OfficePackageSecurityRule.PartCount
                    ? options.MaxPartCount
                    : (double?)null;
                findings.Add(Error(rule, error ?? "The OLE compound package directory could not be inspected.",
                    observedValue: observed, limit: limit));
                return new OfficePackageSecurityReport(bytes.Length, OfficePackageContainerKind.CompoundBinary,
                    0, 0, 0, 0, 0, 0, 0, 0, 0, findings.ToArray());
            }

            int partCount = 0;
            long totalBytes = 0;
            long largestPart = 0;
            int macroCount = 0;
            int embeddedCount = 0;
            int activeXCount = 0;
            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (OfficeCompoundFileEntry entry in entries) {
                if (entry.IsFallback || entry.ObjectType == 5) continue;
                partCount++;
                if (!paths.Add(entry.Path)) {
                    findings.Add(Error(OfficePackageSecurityRule.DuplicatePartName,
                        $"Compound package contains a duplicate or case-ambiguous entry path '{entry.Path}'.",
                        entry.Path));
                }
                if (IsUnsafeCompoundEntry(entry)) {
                    findings.Add(Error(OfficePackageSecurityRule.UnsafePartName,
                        $"Compound package entry name '{entry.Name}' is unsafe.", entry.Path));
                }
                if (!entry.IsStream) continue;

                totalBytes = checked(totalBytes + entry.Size);
                if (entry.Size > largestPart) largestPart = entry.Size;
                if (entry.Size > options.MaxPartUncompressedBytes) {
                    findings.Add(Error(OfficePackageSecurityRule.PartSize,
                        $"Compound stream '{entry.Path}' is {entry.Size} bytes and exceeds the configured maximum of {options.MaxPartUncompressedBytes} bytes.",
                        entry.Path, entry.Size, options.MaxPartUncompressedBytes));
                }
                if (IsCompoundMacroStream(entry.Path)) macroCount++;
                if (IsCompoundEmbeddedStream(entry.Path)) embeddedCount++;
                if (IsCompoundActiveXStream(entry.Path)) activeXCount++;
            }

            AddLimitFinding(findings, OfficePackageSecurityRule.PartCount, partCount,
                options.MaxPartCount, "Compound stream count");
            AddLimitFinding(findings, OfficePackageSecurityRule.TotalUncompressedSize, totalBytes,
                options.MaxTotalUncompressedBytes, "Total compound stream size");
            AddPolicyFinding(findings, options.Macros, OfficePackageSecurityRule.Macros, macroCount,
                "VBA project directory stream");
            AddPolicyFinding(findings, options.EmbeddedPayloads, OfficePackageSecurityRule.EmbeddedPayloads,
                embeddedCount, "embedded ObjectPool stream");
            AddPolicyFinding(findings, options.ActiveX, OfficePackageSecurityRule.ActiveX, activeXCount,
                "ActiveX metadata stream");

            return new OfficePackageSecurityReport(bytes.Length, OfficePackageContainerKind.CompoundBinary,
                partCount, totalBytes, largestPart, 0, macroCount, embeddedCount, activeXCount,
                0, 0, findings.ToArray());
        }

        private static int InspectRelationships(
            ZipArchiveEntry entry,
            string partName,
            ISet<string> macroParts,
            ISet<string> embeddedParts,
            ISet<string> activeXParts,
            ICollection<OfficePackageSecurityFinding> findings) {
            int externalCount = 0;
            try {
                using Stream stream = entry.Open();
                using XmlReader reader = XmlReader.Create(stream, CreateSecureXmlSettings(entry.Length));
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element
                        || !string.Equals(reader.LocalName, "Relationship", StringComparison.Ordinal)) continue;
                    string? targetMode = reader.GetAttribute("TargetMode");
                    if (string.Equals(targetMode, "External", StringComparison.OrdinalIgnoreCase)) {
                        externalCount++;
                        continue;
                    }

                    string? relationshipType = reader.GetAttribute("Type");
                    string? target = reader.GetAttribute("Target");
                    string? targetPart = ResolveRelationshipTarget(partName, target);
                    string classificationKey = targetPart ?? partName + "#" + (target ?? string.Empty);
                    AddRelationshipClassification(relationshipType, classificationKey, macroParts, embeddedParts, activeXParts);
                }
            } catch (Exception exception) when (exception is XmlException || exception is InvalidDataException
                || exception is IOException) {
                findings.Add(Error(OfficePackageSecurityRule.MalformedRelationship,
                    $"Relationship part '{partName}' could not be parsed safely. {exception.Message}", partName));
            }
            return externalCount;
        }

        private static void ValidateOptions(OfficePackageSecurityOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (options.MaxPackageBytes < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxPackageBytes));
            if (options.MaxPartCount < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxPartCount));
            if (options.MaxPartUncompressedBytes < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxPartUncompressedBytes));
            if (options.MaxTotalUncompressedBytes < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxTotalUncompressedBytes));
            if (double.IsNaN(options.MaxCompressionRatio) || double.IsInfinity(options.MaxCompressionRatio)
                || options.MaxCompressionRatio <= 0) {
                throw new ArgumentOutOfRangeException(nameof(options.MaxCompressionRatio));
            }
        }

        private static void AddLimitFinding(ICollection<OfficePackageSecurityFinding> findings,
            OfficePackageSecurityRule rule, long observed, long limit, string label) {
            if (observed <= limit) return;
            findings.Add(Error(rule,
                $"{label} {observed} exceeds the configured maximum of {limit}.",
                observedValue: observed, limit: limit));
        }

        private static void AddPolicyFinding(ICollection<OfficePackageSecurityFinding> findings,
            OfficePackageContentPolicy policy, OfficePackageSecurityRule rule, int count, string contentLabel) {
            if (policy != OfficePackageContentPolicy.Reject || count == 0) return;
            findings.Add(Error(rule,
                $"Package contains {count} {contentLabel}{(count == 1 ? string.Empty : "s")}, which the selected policy rejects.",
                observedValue: count, limit: 0));
        }

        private static OfficePackageSecurityFinding Error(OfficePackageSecurityRule rule, string message,
            string? partName = null, double? observedValue = null, double? limit = null) =>
            new OfficePackageSecurityFinding(OfficePackageSecuritySeverity.Error, rule, message,
                partName, observedValue, limit);

        private static void ThrowFirstError(OfficePackageSecurityReport report) {
            for (int index = 0; index < report.Findings.Count; index++) {
                OfficePackageSecurityFinding finding = report.Findings[index];
                if (finding.Severity == OfficePackageSecuritySeverity.Error) {
                    throw new OfficePackageSecurityException(finding);
                }
            }
        }

        private static bool HasZipSignature(byte[] bytes) => bytes.Length >= 4
            && bytes[0] == 0x50 && bytes[1] == 0x4b
            && ((bytes[2] == 0x03 && bytes[3] == 0x04)
                || (bytes[2] == 0x05 && bytes[3] == 0x06)
                || (bytes[2] == 0x07 && bytes[3] == 0x08));

        private static bool IsDirectory(ZipArchiveEntry entry) => entry.FullName.EndsWith("/", StringComparison.Ordinal)
            || entry.FullName.EndsWith("\\", StringComparison.Ordinal);

        private static string NormalizePartName(string name) => "/" + name.Replace('\\', '/').TrimStart('/');

        private static bool IsUnsafePartName(string name) {
            if (string.IsNullOrWhiteSpace(name) || name[0] == '/' || name[0] == '\\') return true;
            string normalized = name.Replace('\\', '/');
            if (normalized.IndexOf(':') >= 0 || normalized.IndexOf('\0') >= 0) return true;
            string[] segments = normalized.Split('/');
            for (int index = 0; index < segments.Length; index++) {
                if (segments[index] == ".." || segments[index] == ".") return true;
            }
            return false;
        }

        private static bool IsMacroPart(string lowerName) => lowerName.EndsWith("/vbaproject.bin", StringComparison.Ordinal)
            || lowerName.EndsWith("/vbadata.xml", StringComparison.Ordinal);

        private static bool IsEmbeddedPart(string lowerName) => lowerName.IndexOf("/embeddings/", StringComparison.Ordinal) >= 0;

        private static bool IsActiveXPart(string lowerName) => lowerName.IndexOf("/activex/", StringComparison.Ordinal) >= 0;

        private static bool IsDigitalSignaturePart(string lowerName) => lowerName.IndexOf("/_xmlsignatures/", StringComparison.Ordinal) >= 0;

        private static bool IsUnsafeCompoundEntry(OfficeCompoundFileEntry entry) =>
            string.IsNullOrWhiteSpace(entry.Name)
            || entry.Name == "."
            || entry.Name == ".."
            || entry.Name.IndexOf('/') >= 0
            || entry.Name.IndexOf('\\') >= 0
            || entry.Name.IndexOf('\0') >= 0;

        private static bool IsCompoundMacroStream(string path) {
            string normalized = "/" + path.Replace('\\', '/').Trim('/').ToLowerInvariant() + "/";
            return normalized.IndexOf("/vba/dir/", StringComparison.Ordinal) >= 0;
        }

        private static bool IsCompoundEmbeddedStream(string path) {
            string normalized = "/" + path.Replace('\\', '/').Trim('/').ToLowerInvariant() + "/";
            return normalized.IndexOf("/objectpool/", StringComparison.Ordinal) >= 0
                || normalized.IndexOf("/mbd", StringComparison.Ordinal) >= 0;
        }

        private static bool IsCompoundActiveXStream(string path) {
            string normalized = "/" + path.Replace('\\', '/').Trim('/').ToLowerInvariant() + "/";
            return normalized.IndexOf("/ctls/", StringComparison.Ordinal) >= 0
                || normalized.IndexOf("/ocxname/", StringComparison.Ordinal) >= 0;
        }

        private static string FormatRatio(double value) => double.IsPositiveInfinity(value)
            ? "infinite"
            : value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
    }
}

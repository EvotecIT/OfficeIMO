using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing {
    /// <summary>Controls whether a class of active or externally linked package content is accepted.</summary>
    public enum OfficePackageContentPolicy {
        /// <summary>The content is inventoried but is not rejected.</summary>
        Allow,
        /// <summary>The package is rejected when the content is present.</summary>
        Reject
    }

    /// <summary>Identifies the physical container inspected by the package security engine.</summary>
    public enum OfficePackageContainerKind {
        /// <summary>The bytes do not identify a supported package container.</summary>
        Unknown,
        /// <summary>The artifact is an Open XML ZIP package.</summary>
        OpenXml,
        /// <summary>The artifact is an OLE compound binary container.</summary>
        CompoundBinary
    }

    /// <summary>Identifies a package security rule.</summary>
    public enum OfficePackageSecurityRule {
        /// <summary>The complete source exceeds the configured byte limit.</summary>
        PackageSize,
        /// <summary>The package contains more parts than allowed.</summary>
        PartCount,
        /// <summary>An individual uncompressed part exceeds the configured byte limit.</summary>
        PartSize,
        /// <summary>The sum of uncompressed part sizes exceeds the configured byte limit.</summary>
        TotalUncompressedSize,
        /// <summary>A ZIP part exceeds the configured compression ratio.</summary>
        CompressionRatio,
        /// <summary>The ZIP contains ambiguous duplicate part names.</summary>
        DuplicatePartName,
        /// <summary>A ZIP part name is absolute, traverses a parent, or otherwise has an unsafe shape.</summary>
        UnsafePartName,
        /// <summary>A relationship part could not be parsed safely.</summary>
        MalformedRelationship,
        /// <summary>The package contains a VBA project while macro content is rejected.</summary>
        Macros,
        /// <summary>The package contains an embedded payload while embedded content is rejected.</summary>
        EmbeddedPayloads,
        /// <summary>The package contains ActiveX content while ActiveX is rejected.</summary>
        ActiveX,
        /// <summary>The package contains external relationships while external targets are rejected.</summary>
        ExternalRelationships,
        /// <summary>The ZIP or compound container is malformed.</summary>
        MalformedPackage
    }

    /// <summary>Severity of a package security finding.</summary>
    public enum OfficePackageSecuritySeverity {
        /// <summary>The finding is informational.</summary>
        Information,
        /// <summary>The finding deserves caller attention but does not make validation fail.</summary>
        Warning,
        /// <summary>The finding violates the selected policy and makes validation fail.</summary>
        Error
    }

    /// <summary>Resource limits and active-content policy applied before Word or Excel opens a package.</summary>
    public sealed class OfficePackageSecurityOptions {
        /// <summary>Maximum complete source size. Defaults to 512 MiB.</summary>
        public long MaxPackageBytes { get; set; } = 512L * 1024L * 1024L;

        /// <summary>Maximum number of Open XML parts or compound directory entries. Defaults to 10,000.</summary>
        public int MaxPartCount { get; set; } = 10_000;

        /// <summary>Maximum uncompressed size of one part. Defaults to 256 MiB.</summary>
        public long MaxPartUncompressedBytes { get; set; } = 256L * 1024L * 1024L;

        /// <summary>Maximum aggregate uncompressed part size. Defaults to 1 GiB.</summary>
        public long MaxTotalUncompressedBytes { get; set; } = 1024L * 1024L * 1024L;

        /// <summary>Maximum uncompressed-to-compressed ratio for a non-empty ZIP part. Defaults to 1,000.</summary>
        public double MaxCompressionRatio { get; set; } = 1000D;

        /// <summary>Policy for VBA projects.</summary>
        public OfficePackageContentPolicy Macros { get; set; } = OfficePackageContentPolicy.Allow;

        /// <summary>Policy for embedded OLE objects and package payloads.</summary>
        public OfficePackageContentPolicy EmbeddedPayloads { get; set; } = OfficePackageContentPolicy.Allow;

        /// <summary>Policy for ActiveX controls.</summary>
        public OfficePackageContentPolicy ActiveX { get; set; } = OfficePackageContentPolicy.Allow;

        /// <summary>Policy for relationships whose target mode is external.</summary>
        public OfficePackageContentPolicy ExternalRelationships { get; set; } = OfficePackageContentPolicy.Allow;

        /// <summary>Creates structural package-bomb limits while retaining active-content compatibility.</summary>
        public static OfficePackageSecurityOptions SecureDefaults => new OfficePackageSecurityOptions();

        /// <summary>Creates structural limits and rejects active, embedded, and externally linked content.</summary>
        public static OfficePackageSecurityOptions UntrustedDefaults => new OfficePackageSecurityOptions {
            Macros = OfficePackageContentPolicy.Reject,
            EmbeddedPayloads = OfficePackageContentPolicy.Reject,
            ActiveX = OfficePackageContentPolicy.Reject,
            ExternalRelationships = OfficePackageContentPolicy.Reject
        };
    }

    /// <summary>Describes one package security observation or policy violation.</summary>
    public sealed class OfficePackageSecurityFinding {
        internal OfficePackageSecurityFinding(OfficePackageSecuritySeverity severity,
            OfficePackageSecurityRule rule, string message, string? partName = null,
            double? observedValue = null, double? limit = null) {
            Severity = severity;
            Rule = rule;
            Message = message;
            PartName = partName;
            ObservedValue = observedValue;
            Limit = limit;
        }

        /// <summary>Finding severity.</summary>
        public OfficePackageSecuritySeverity Severity { get; }

        /// <summary>Rule that produced the finding.</summary>
        public OfficePackageSecurityRule Rule { get; }

        /// <summary>Human-readable explanation.</summary>
        public string Message { get; }

        /// <summary>Package-local part name, when the finding is part-specific.</summary>
        public string? PartName { get; }

        /// <summary>Observed numeric value, when relevant.</summary>
        public double? ObservedValue { get; }

        /// <summary>Configured limit, when relevant.</summary>
        public double? Limit { get; }
    }

    /// <summary>Inventory and validation findings produced without executing package content.</summary>
    public sealed class OfficePackageSecurityReport {
        internal OfficePackageSecurityReport(long packageBytes, OfficePackageContainerKind containerKind,
            int partCount, long totalUncompressedBytes, long largestPartBytes, double highestCompressionRatio,
            int macroPartCount, int embeddedPayloadPartCount, int activeXPartCount,
            int externalRelationshipCount, int digitalSignaturePartCount,
            IReadOnlyList<OfficePackageSecurityFinding> findings) {
            PackageBytes = packageBytes;
            ContainerKind = containerKind;
            PartCount = partCount;
            TotalUncompressedBytes = totalUncompressedBytes;
            LargestPartBytes = largestPartBytes;
            HighestCompressionRatio = highestCompressionRatio;
            MacroPartCount = macroPartCount;
            EmbeddedPayloadPartCount = embeddedPayloadPartCount;
            ActiveXPartCount = activeXPartCount;
            ExternalRelationshipCount = externalRelationshipCount;
            DigitalSignaturePartCount = digitalSignaturePartCount;
            Findings = findings;
        }

        /// <summary>Complete source size in bytes.</summary>
        public long PackageBytes { get; }

        /// <summary>Detected physical container.</summary>
        public OfficePackageContainerKind ContainerKind { get; }

        /// <summary>Number of non-directory ZIP parts or non-root compound directory entries.</summary>
        public int PartCount { get; }

        /// <summary>Aggregate uncompressed ZIP part size.</summary>
        public long TotalUncompressedBytes { get; }

        /// <summary>Largest uncompressed part size.</summary>
        public long LargestPartBytes { get; }

        /// <summary>Highest observed ZIP compression ratio.</summary>
        public double HighestCompressionRatio { get; }

        /// <summary>Number of VBA project parts.</summary>
        public int MacroPartCount { get; }

        /// <summary>Number of embedded package or OLE payload parts.</summary>
        public int EmbeddedPayloadPartCount { get; }

        /// <summary>Number of ActiveX parts.</summary>
        public int ActiveXPartCount { get; }

        /// <summary>Number of relationships with an external target.</summary>
        public int ExternalRelationshipCount { get; }

        /// <summary>Number of digital-signature package parts.</summary>
        public int DigitalSignaturePartCount { get; }

        /// <summary>Findings produced by structural validation and the selected content policy.</summary>
        public IReadOnlyList<OfficePackageSecurityFinding> Findings { get; }

        /// <summary>True when the report contains no error findings.</summary>
        public bool IsValid {
            get {
                for (int index = 0; index < Findings.Count; index++) {
                    if (Findings[index].Severity == OfficePackageSecuritySeverity.Error) return false;
                }
                return true;
            }
        }
    }

    /// <summary>Signals a package resource-limit or content-policy violation.</summary>
    public sealed class OfficePackageSecurityException : IOException {
        internal OfficePackageSecurityException(OfficePackageSecurityFinding finding)
            : base(finding == null ? "Package security validation failed." : finding.Message) {
            if (finding == null) throw new ArgumentNullException(nameof(finding));
            Rule = finding.Rule;
            PartName = finding.PartName;
            ObservedValue = finding.ObservedValue;
            Limit = finding.Limit;
        }

        internal OfficePackageSecurityException(OfficePackageSecurityRule rule, string message,
            double? observedValue = null, double? limit = null)
            : base(message) {
            Rule = rule;
            ObservedValue = observedValue;
            Limit = limit;
        }

        /// <summary>Rule that rejected the package.</summary>
        public OfficePackageSecurityRule Rule { get; }

        /// <summary>Package-local part name, when the rejection is part-specific.</summary>
        public string? PartName { get; }

        /// <summary>Observed numeric value, when relevant.</summary>
        public double? ObservedValue { get; }

        /// <summary>Configured limit, when relevant.</summary>
        public double? Limit { get; }
    }
}

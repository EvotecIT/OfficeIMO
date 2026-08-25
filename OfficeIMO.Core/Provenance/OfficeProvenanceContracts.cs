using System;
using System.Collections.Generic;

namespace OfficeIMO.Provenance;

/// <summary>Identifies the asset container inspected for provenance.</summary>
public enum OfficeProvenanceAssetFormat {
    /// <summary>The asset format could not be identified.</summary>
    Unknown,
    /// <summary>JPEG image data.</summary>
    Jpeg,
    /// <summary>PNG image data.</summary>
    Png,
    /// <summary>WebP image data in a RIFF container.</summary>
    Webp,
    /// <summary>GIF image data.</summary>
    Gif,
    /// <summary>TIFF or TIFF-derived image data.</summary>
    Tiff,
    /// <summary>SVG XML image data.</summary>
    Svg,
    /// <summary>A ZIP-based package such as OOXML, OpenDocument, or EPUB.</summary>
    ZipPackage,
    /// <summary>Structured text carrying an ASCII-armoured manifest reference.</summary>
    StructuredText,
    /// <summary>Unstructured text carrying a variation-selector manifest wrapper.</summary>
    UnstructuredText,
    /// <summary>An HTML document.</summary>
    Html,
    /// <summary>A PDF document.</summary>
    Pdf
}

/// <summary>Identifies a standards-defined provenance carrier.</summary>
public enum OfficeProvenanceCarrierKind {
    /// <summary>An embedded C2PA Manifest Store.</summary>
    C2paManifest,
    /// <summary>A reference to an external C2PA Manifest Store.</summary>
    C2paExternalManifest,
    /// <summary>An IPTC Digital Source Type declaration carried by XMP.</summary>
    IptcDigitalSourceType
}

/// <summary>Describes the origin declared by an IPTC Digital Source Type value.</summary>
public enum OfficeProvenanceDigitalSourceKind {
    /// <summary>The value is not recognized by this version of OfficeIMO.</summary>
    Unknown,
    /// <summary>Content captured from the physical world.</summary>
    DigitalCapture,
    /// <summary>Content produced by a non-trained algorithm.</summary>
    AlgorithmicMedia,
    /// <summary>Content produced by a trained generative model.</summary>
    TrainedAlgorithmicMedia,
    /// <summary>A composite containing trained-algorithmic media.</summary>
    CompositeWithTrainedAlgorithmicMedia,
    /// <summary>A composite assembled from captured media.</summary>
    CompositeCapture,
    /// <summary>A digital source type explicitly outside the preceding categories.</summary>
    Other
}

/// <summary>One provenance signal discovered in an asset.</summary>
public sealed class OfficeProvenanceEvidence {
    /// <summary>Creates provenance evidence.</summary>
    public OfficeProvenanceEvidence(
        OfficeProvenanceCarrierKind carrier,
        string location,
        bool isStructurallyValid,
        long payloadLength = 0,
        string? value = null,
        OfficeProvenanceDigitalSourceKind digitalSourceKind = OfficeProvenanceDigitalSourceKind.Unknown) {
        if (string.IsNullOrWhiteSpace(location)) throw new ArgumentException("A carrier location is required.", nameof(location));
        if (payloadLength < 0) throw new ArgumentOutOfRangeException(nameof(payloadLength));
        Carrier = carrier;
        Location = location;
        IsStructurallyValid = isStructurallyValid;
        PayloadLength = payloadLength;
        Value = value;
        DigitalSourceKind = digitalSourceKind;
    }

    /// <summary>Gets the carrier kind.</summary>
    public OfficeProvenanceCarrierKind Carrier { get; }
    /// <summary>Gets a format-native carrier location.</summary>
    public string Location { get; }
    /// <summary>Gets whether the carrier has the required structural shape.</summary>
    public bool IsStructurallyValid { get; }
    /// <summary>Gets the embedded payload length, when known.</summary>
    public long PayloadLength { get; }
    /// <summary>Gets the external URI or declared source-type value, when applicable.</summary>
    public string? Value { get; }
    /// <summary>Gets the classified digital source type, when applicable.</summary>
    public OfficeProvenanceDigitalSourceKind DigitalSourceKind { get; }
}

/// <summary>Structural provenance inspection of one asset.</summary>
public sealed class OfficeProvenanceReport {
    /// <summary>Creates an inspection report.</summary>
    public OfficeProvenanceReport(
        OfficeProvenanceAssetFormat format,
        IReadOnlyList<OfficeProvenanceEvidence> evidence,
        IReadOnlyList<string>? diagnostics = null) {
        Format = format;
        Evidence = new List<OfficeProvenanceEvidence>(evidence ?? throw new ArgumentNullException(nameof(evidence))).AsReadOnly();
        Diagnostics = new List<string>(diagnostics ?? Array.Empty<string>()).AsReadOnly();
    }

    /// <summary>Gets the identified asset format.</summary>
    public OfficeProvenanceAssetFormat Format { get; }
    /// <summary>Gets discovered provenance evidence in source order.</summary>
    public IReadOnlyList<OfficeProvenanceEvidence> Evidence { get; }
    /// <summary>Gets structural diagnostics that did not prevent inspection.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
    /// <summary>Gets whether an embedded C2PA carrier was discovered.</summary>
    public bool HasC2paManifest {
        get {
            foreach (OfficeProvenanceEvidence item in Evidence) {
                if (item.Carrier == OfficeProvenanceCarrierKind.C2paManifest) return true;
            }
            return false;
        }
    }
    /// <summary>Gets whether an external C2PA reference was discovered.</summary>
    public bool HasExternalC2paManifest {
        get {
            foreach (OfficeProvenanceEvidence item in Evidence) {
                if (item.Carrier == OfficeProvenanceCarrierKind.C2paExternalManifest) return true;
            }
            return false;
        }
    }
    /// <summary>Gets whether a trained-algorithmic source declaration was discovered.</summary>
    public bool HasGenerativeAiDeclaration {
        get {
            foreach (OfficeProvenanceEvidence item in Evidence) {
                if (item.DigitalSourceKind == OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia ||
                    item.DigitalSourceKind == OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia) return true;
            }
            return false;
        }
    }
}

/// <summary>Bounds structural provenance inspection and removal.</summary>
public sealed class OfficeProvenanceOptions {
    /// <summary>Maximum encoded asset bytes accepted. Defaults to 256 MiB.</summary>
    public long MaxAssetBytes { get; set; } = 256L * 1024L * 1024L;
    /// <summary>Maximum single manifest-store bytes accepted. Defaults to 64 MiB.</summary>
    public long MaxManifestBytes { get; set; } = 64L * 1024L * 1024L;
    /// <summary>Maximum carriers accepted in one asset. Defaults to 128.</summary>
    public int MaxCarriers { get; set; } = 128;
    /// <summary>Maximum structural entries or materialized XML nodes accepted in a container. Defaults to 65,536.</summary>
    public int MaxContainerEntries { get; set; } = 65536;
    /// <summary>Maximum cumulative expanded bytes copied while rewriting a container. Defaults to 1 GiB.</summary>
    public long MaxExpandedContainerBytes { get; set; } = 1024L * 1024L * 1024L;
    /// <summary>Whether supported image assets inside ZIP-based documents are inspected. Defaults to true.</summary>
    public bool ProcessEmbeddedAssets { get; set; } = true;
    /// <summary>Maximum supported embedded assets inspected in one container. Defaults to 4,096.</summary>
    public int MaxEmbeddedAssets { get; set; } = 4096;
}

/// <summary>Controls selective provenance removal.</summary>
public sealed class OfficeProvenanceRemovalOptions {
    /// <summary>Whether embedded C2PA Manifest Stores are removed. Defaults to true.</summary>
    public bool RemoveC2paManifests { get; set; } = true;
    /// <summary>Whether external C2PA references are removed. Defaults to true.</summary>
    public bool RemoveExternalC2paReferences { get; set; } = true;
    /// <summary>Whether AI-specific IPTC Digital Source Type declarations are removed. Defaults to true.</summary>
    public bool RemoveAiSourceMetadata { get; set; } = true;
    /// <summary>Whether malformed carriers are preserved instead of removed. Defaults to true.</summary>
    public bool RequireStructurallyValidCarrier { get; set; } = true;
    /// <summary>Controls how package signatures invalidated by removal are handled. Defaults to blocking mutation.</summary>
    public OfficeIMO.OfficeSignatureMutationPolicy SignatureMutationPolicy { get; set; } = OfficeIMO.OfficeSignatureMutationPolicy.BlockSave;
    /// <summary>Whether document integrations inspect and sanitize supported embedded assets. Defaults to true.</summary>
    public bool ProcessEmbeddedAssets { get; set; } = true;
    /// <summary>Maximum embedded assets processed by one document operation. Defaults to 4096.</summary>
    public int MaxEmbeddedAssets { get; set; } = 4096;
    /// <summary>Inspection and removal resource limits.</summary>
    public OfficeProvenanceOptions Limits { get; } = new OfficeProvenanceOptions();
}

/// <summary>One format-native provenance mutation.</summary>
public sealed class OfficeProvenanceChange {
    /// <summary>Creates a mutation descriptor.</summary>
    public OfficeProvenanceChange(OfficeProvenanceCarrierKind carrier, string location, long removedBytes) {
        Carrier = carrier;
        Location = location ?? throw new ArgumentNullException(nameof(location));
        RemovedBytes = removedBytes;
    }
    /// <summary>Gets the removed carrier kind.</summary>
    public OfficeProvenanceCarrierKind Carrier { get; }
    /// <summary>Gets the format-native location that changed.</summary>
    public string Location { get; }
    /// <summary>Gets the number of physical bytes removed, or zero when a container was rewritten in place.</summary>
    public long RemovedBytes { get; }
}

/// <summary>Result of a provenance-removal operation.</summary>
public sealed class OfficeProvenanceRemovalResult {
    private readonly byte[] _data;

    /// <summary>Creates a provenance-removal result.</summary>
    public OfficeProvenanceRemovalResult(
        byte[] data,
        OfficeProvenanceReport before,
        OfficeProvenanceReport after,
        IReadOnlyList<OfficeProvenanceChange> changes,
        bool wasReserialized)
        : this(data, before, after, changes, wasReserialized, wereInvalidatedSignaturesRemoved: false) {
    }

    /// <summary>Creates a provenance-removal result and records whether invalidated signatures were removed.</summary>
    public OfficeProvenanceRemovalResult(
        byte[] data,
        OfficeProvenanceReport before,
        OfficeProvenanceReport after,
        IReadOnlyList<OfficeProvenanceChange> changes,
        bool wasReserialized,
        bool wereInvalidatedSignaturesRemoved)
        : this(
            data,
            before,
            after,
            changes,
            wasReserialized,
            wereInvalidatedSignaturesRemoved,
            takeOwnership: false) {
    }

    private OfficeProvenanceRemovalResult(
        byte[] data,
        OfficeProvenanceReport before,
        OfficeProvenanceReport after,
        IReadOnlyList<OfficeProvenanceChange> changes,
        bool wasReserialized,
        bool wereInvalidatedSignaturesRemoved,
        bool takeOwnership) {
        _data = takeOwnership
            ? data ?? throw new ArgumentNullException(nameof(data))
            : (byte[])(data ?? throw new ArgumentNullException(nameof(data))).Clone();
        Before = before ?? throw new ArgumentNullException(nameof(before));
        After = after ?? throw new ArgumentNullException(nameof(after));
        Changes = new List<OfficeProvenanceChange>(changes ?? throw new ArgumentNullException(nameof(changes))).AsReadOnly();
        WasReserialized = wasReserialized;
        WereInvalidatedSignaturesRemoved = wereInvalidatedSignaturesRemoved;
    }

    internal static OfficeProvenanceRemovalResult CreateOwned(
        byte[] data,
        OfficeProvenanceReport before,
        OfficeProvenanceReport after,
        IReadOnlyList<OfficeProvenanceChange> changes,
        bool wasReserialized,
        bool wereInvalidatedSignaturesRemoved = false) =>
        new OfficeProvenanceRemovalResult(
            data,
            before,
            after,
            changes,
            wasReserialized,
            wereInvalidatedSignaturesRemoved,
            takeOwnership: true);

    /// <summary>Gets the inspection before removal.</summary>
    public OfficeProvenanceReport Before { get; }
    /// <summary>Gets the inspection after removal.</summary>
    public OfficeProvenanceReport After { get; }
    /// <summary>Gets format-native changes in source order.</summary>
    public IReadOnlyList<OfficeProvenanceChange> Changes { get; }
    /// <summary>Gets whether any carrier changed.</summary>
    public bool WasChanged => Changes.Count != 0;
    /// <summary>Gets whether the container was serialized rather than copied byte-for-byte around removed carriers.</summary>
    public bool WasReserialized { get; }
    /// <summary>Gets whether an owning document API removed signatures that the provenance mutation invalidated.</summary>
    public bool WereInvalidatedSignaturesRemoved { get; }
    /// <summary>Returns an owned copy of the resulting asset.</summary>
    public byte[] ToArray() => (byte[])_data.Clone();
}

/// <summary>Outcome assigned by an optional provenance-verification provider.</summary>
public enum OfficeProvenanceVerificationStatus {
    /// <summary>No provenance carrier was present.</summary>
    NotPresent,
    /// <summary>The provider verified content binding and signature mathematics.</summary>
    Valid,
    /// <summary>The provider found a failed binding, signature, or malformed claim.</summary>
    Invalid,
    /// <summary>The carrier was parsed but trust could not be established.</summary>
    Untrusted,
    /// <summary>The provider could not reach a definitive outcome.</summary>
    Indeterminate,
    /// <summary>The configured provider was unavailable.</summary>
    ProviderUnavailable,
    /// <summary>The provider failed before producing trustworthy evidence.</summary>
    Error
}

/// <summary>Neutral options for optional cryptographic provenance verification.</summary>
public sealed class OfficeProvenanceVerificationOptions {
    /// <summary>Maximum provider runtime. Defaults to 30 seconds.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(30);
    /// <summary>Maximum provider report bytes accepted. Defaults to 8 MiB.</summary>
    public long MaxReportBytes { get; set; } = 8L * 1024L * 1024L;
    /// <summary>Whether the verifier may resolve remote manifests or trust material. Defaults to false.</summary>
    public bool AllowNetworkAccess { get; set; }
    /// <summary>Whether the bounded provider JSON is returned with the normalized result. Defaults to false.</summary>
    public bool IncludeRawReport { get; set; }
    /// <summary>Optional local PEM trust-anchor list.</summary>
    public string? TrustAnchorsPath { get; set; }
    /// <summary>Optional local PEM allowed-certificate list.</summary>
    public string? AllowedListPath { get; set; }
    /// <summary>Optional local C2PA trust configuration.</summary>
    public string? TrustConfigurationPath { get; set; }
}

/// <summary>Cryptographic and trust evidence returned by an optional provider.</summary>
public sealed class OfficeProvenanceVerificationResult {
    /// <summary>Creates provider verification evidence.</summary>
    public OfficeProvenanceVerificationResult(
        OfficeProvenanceVerificationStatus status,
        string providerName,
        IReadOnlyList<string> findings,
        string? rawReport = null) {
        Status = status;
        ProviderName = providerName ?? throw new ArgumentNullException(nameof(providerName));
        Findings = new List<string>(findings ?? throw new ArgumentNullException(nameof(findings))).AsReadOnly();
        RawReport = rawReport;
    }
    /// <summary>Gets the combined verification outcome.</summary>
    public OfficeProvenanceVerificationStatus Status { get; }
    /// <summary>Gets the verifier implementation name.</summary>
    public string ProviderName { get; }
    /// <summary>Gets normalized validation findings.</summary>
    public IReadOnlyList<string> Findings { get; }
    /// <summary>Gets the bounded provider report, when requested and available.</summary>
    public string? RawReport { get; }
}

/// <summary>Optional provider contract for cryptographic provenance verification.</summary>
public interface IOfficeProvenanceVerifier {
    /// <summary>Gets the provider name.</summary>
    string Name { get; }
    /// <summary>Verifies provenance carried by an asset file.</summary>
    OfficeProvenanceVerificationResult Verify(string filePath, OfficeProvenanceVerificationOptions? options = null);
}

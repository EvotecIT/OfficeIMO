using System.Threading;

namespace OfficeIMO.Epub;

/// <summary>The package generation used for newly authored publications.</summary>
public enum EpubVersion {
    /// <summary>EPUB 2.0.1 package and NCX navigation.</summary>
    Epub2,
    /// <summary>EPUB 3 package and XHTML navigation.</summary>
    Epub3
}

/// <summary>Bounds the complete package retained for editing, independently of chapter extraction.</summary>
public sealed class EpubPublicationLoadOptions {
    /// <summary>Maximum compressed input bytes.</summary>
    public long MaxInputBytes { get; set; } = 128L * 1024 * 1024;
    /// <summary>Maximum total expanded bytes retained, including unknown entries and current selected package XML.</summary>
    public long MaxExpandedBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum individual entry bytes.</summary>
    public long MaxEntryBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum package/container XML bytes, also enforced on edits to the selected package.</summary>
    public long MaxMetadataBytes { get; set; } = 4L * 1024 * 1024;
    /// <summary>Maximum ZIP entries, including directories and the selected package XML; also enforced when adding resources.</summary>
    public int MaxEntries { get; set; } = 10_000;
}

/// <summary>Bounds serialized output and controls explicit signature invalidation.</summary>
public sealed class EpubWriteOptions {
    /// <summary>Deflates ordinary entries when true. The leading mimetype entry is always stored.</summary>
    public bool CompressEntries { get; set; } = true;
    /// <summary>Maximum compressed output bytes, checked while producing the ZIP.</summary>
    public long MaxOutputBytes { get; set; } = 128L * 1024 * 1024;
    /// <summary>Maximum total expanded output bytes.</summary>
    public long MaxExpandedBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum number of output entries.</summary>
    public int MaxEntries { get; set; } = 10_000;
    /// <summary>Explicitly permits removing META-INF/signatures.xml when edits invalidate it.</summary>
    public bool RemoveInvalidatedSignatures { get; set; }
    /// <summary>Optional UTC modification time. Otherwise a stable time is captured when edits begin.</summary>
    public DateTimeOffset? ModifiedAt { get; set; }
}

/// <summary>Preservation and omission evidence for one writer operation; this is not a conformance certificate.</summary>
public sealed class EpubWriteReport : IOfficeConversionReport {
    internal EpubWriteReport(bool usedOriginal, IEnumerable<string> preserved, IEnumerable<string> regenerated, IEnumerable<string> removed,
        IEnumerable<OfficeConversionFidelityDiagnostic> diagnostics, IReadOnlyDictionary<string, string>? renamed = null, IReadOnlyDictionary<string, string>? merged = null) {
        UsedOriginalPackage = usedOriginal;
        PreservedEntries = Array.AsReadOnly(preserved.ToArray());
        RegeneratedEntries = Array.AsReadOnly(regenerated.ToArray());
        RemovedEntries = Array.AsReadOnly(removed.ToArray());
        FidelityDiagnostics = Array.AsReadOnly(diagnostics.ToArray());
        RenamedEntries = new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>((renamed ?? new Dictionary<string, string>()).ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal));
        MergedEntries = new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>((merged ?? new Dictionary<string, string>()).ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal));
    }
    /// <summary>Whether the exact original compressed package was returned.</summary>
    public bool UsedOriginalPackage { get; }
    /// <summary>Entry payloads retained without byte changes.</summary>
    public IReadOnlyList<string> PreservedEntries { get; }
    /// <summary>Entry payloads created or changed by authoring or serialization.</summary>
    public IReadOnlyList<string> RegeneratedEntries { get; }
    /// <summary>Original paths absent from output, including renamed entries, explicit removal and signature policy.</summary>
    public IReadOnlyList<string> RemovedEntries { get; }
    /// <summary>Original entry paths moved to current paths by resource renaming, rather than omitted.</summary>
    public IReadOnlyDictionary<string, string> RenamedEntries { get; }
    /// <summary>Original resource paths whose content was consolidated into current paths by chapter merging.</summary>
    public IReadOnlyDictionary<string, string> MergedEntries { get; }
    /// <inheritdoc />
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }
    /// <inheritdoc />
    public bool HasLoss => FidelityDiagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new OfficeConversionException("EPUB writing reported content loss.", this);
    }
}

/// <summary>Serialized EPUB bytes and their operation-level preservation report.</summary>
public sealed class EpubWriteResult : OfficeConversionResult<byte[], EpubWriteReport> {
    internal EpubWriteResult(byte[] bytes, EpubWriteReport report) : base(bytes, report) { }
    /// <summary>Complete serialized artifact.</summary>
    public byte[] Bytes => Value;
}

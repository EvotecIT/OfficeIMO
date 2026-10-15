using System;

namespace OfficeIMO;

/// <summary>Bounds one inert embedded-payload inventory and its selected reads. Native document input and parser limits apply separately.</summary>
public sealed class OfficeEmbeddedPayloadInspectionOptions {
    /// <summary>Maximum payload descriptors admitted. Default: 512.</summary>
    public int MaxPayloads { get; set; } = 512;
    /// <summary>Maximum decoded bytes in one payload. Default: 32 MiB.</summary>
    public long MaxPayloadBytes { get; set; } = 32L * 1024 * 1024;
    /// <summary>Maximum sum of known declared and subsequently observed lengths, counting each inventory descriptor once. Distinct owner relationships remain distinct descriptors. Default: 128 MiB.</summary>
    public long MaxTotalPayloadBytes { get; set; } = 128L * 1024 * 1024;
    /// <summary>Maximum actual decoded bytes consumed across hash scans and selected reads, including repeated reads. Default: 256 MiB.</summary>
    public long MaxReadBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Whether to scan every payload for SHA-256 during inventory. Default: false; unknown lengths remain explicit without an unrequested scan.</summary>
    public bool IncludeSha256 { get; set; }

    internal OfficeEmbeddedPayloadInspectionOptions Snapshot() {
        if (MaxPayloads < 0) throw new ArgumentOutOfRangeException(nameof(MaxPayloads));
        if (MaxPayloadBytes < 0) throw new ArgumentOutOfRangeException(nameof(MaxPayloadBytes));
        if (MaxTotalPayloadBytes < 0) throw new ArgumentOutOfRangeException(nameof(MaxTotalPayloadBytes));
        if (MaxReadBytes < 0) throw new ArgumentOutOfRangeException(nameof(MaxReadBytes));
        return new OfficeEmbeddedPayloadInspectionOptions {
            MaxPayloads = MaxPayloads, MaxPayloadBytes = MaxPayloadBytes,
            MaxTotalPayloadBytes = MaxTotalPayloadBytes, MaxReadBytes = MaxReadBytes, IncludeSha256 = IncludeSha256
        };
    }
}

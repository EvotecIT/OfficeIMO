using System.Collections.Generic;

namespace OfficeIMO.Archives;

/// <summary>The result of a complete member verification attempt.</summary>
public enum ArchiveVerificationStatus {
    /// <summary>Complete member read and all applicable integrity checks succeeded.</summary>
    Verified,
    /// <summary>The declared source or compression is unsupported.</summary>
    Unsupported,
    /// <summary>The member failed an integrity or consistency check.</summary>
    Invalid
}

/// <summary>A structured outcome retaining the authored source identity.</summary>
public sealed class ArchiveMemberVerification {
    internal ArchiveMemberVerification(ArchiveInspectionEntry entry, ArchiveVerificationStatus status, string? code, string? message) {
        Entry = entry; Status = status; Code = code; Message = message;
    }
    /// <summary>The original immutable inventory entry.</summary>
    public ArchiveInspectionEntry Entry { get; }
    /// <summary>Verification outcome; CAB verification covers only declared nonzero checksums and structural integrity.</summary>
    public ArchiveVerificationStatus Status { get; }
    /// <summary>Stable diagnostic code, or null on success.</summary>
    public string? Code { get; }
    /// <summary>Diagnostic detail, or null on success. Do not classify failures by this text.</summary>
    public string? Message { get; }
}

/// <summary>Detached outcomes and cumulative resource usage. Cancellation and resource failures throw instead of returning a partial report.</summary>
public sealed class ArchiveVerificationResult {
    internal ArchiveVerificationResult(IReadOnlyList<ArchiveMemberVerification> members, ArchiveResourceUsage usage) {
        Members = members; ResourceUsage = usage;
    }
    /// <summary>One result per requested source ordinal, in request order.</summary>
    public IReadOnlyList<ArchiveMemberVerification> Members { get; }
    /// <summary>Cumulative charges at completion.</summary>
    public ArchiveResourceUsage ResourceUsage { get; }
}

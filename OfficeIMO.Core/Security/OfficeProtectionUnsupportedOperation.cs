using System;

namespace OfficeIMO.Security;

/// <summary>Identifies one operation in a protected-content capability row.</summary>
public enum OfficeProtectionOperation {
    /// <summary>Structural or semantic inspection.</summary>
    Inspect,
    /// <summary>Decryption, deobfuscation, or protected open.</summary>
    Open,
    /// <summary>Protection or signature creation.</summary>
    Create,
    /// <summary>Cryptographic or structural validation.</summary>
    Validate,
    /// <summary>Safe mutation or round-trip.</summary>
    Mutate,
    /// <summary>Explicit protection removal.</summary>
    Remove
}

/// <summary>Explains why a protected-content operation is not supported.</summary>
public enum OfficeProtectionUnsupportedDisposition {
    /// <summary>The missing operation is owned by an open roadmap outcome.</summary>
    RoadmapTracked,
    /// <summary>The operation is outside the deliberately supported format or provider boundary.</summary>
    IntentionalBoundary
}

/// <summary>Machine-readable disposition for one unsupported protected-content operation.</summary>
public sealed class OfficeProtectionUnsupportedOperation {
    /// <summary>Creates an unsupported-operation disposition.</summary>
    public OfficeProtectionUnsupportedOperation(
        OfficeProtectionOperation operation,
        OfficeProtectionUnsupportedDisposition disposition,
        string rationale,
        string? roadmapReference = null) {
        if (!Enum.IsDefined(typeof(OfficeProtectionOperation), operation)) {
            throw new ArgumentOutOfRangeException(nameof(operation));
        }
        if (!Enum.IsDefined(typeof(OfficeProtectionUnsupportedDisposition), disposition)) {
            throw new ArgumentOutOfRangeException(nameof(disposition));
        }
        if (string.IsNullOrWhiteSpace(rationale)) {
            throw new ArgumentException("Unsupported-operation rationale cannot be empty.", nameof(rationale));
        }
        string? normalizedReference = string.IsNullOrWhiteSpace(roadmapReference)
            ? null
            : roadmapReference!.Trim();
        if (disposition == OfficeProtectionUnsupportedDisposition.RoadmapTracked && normalizedReference == null) {
            throw new ArgumentException("A roadmap-tracked operation requires a roadmap reference.", nameof(roadmapReference));
        }
        if (disposition == OfficeProtectionUnsupportedDisposition.IntentionalBoundary && normalizedReference != null) {
            throw new ArgumentException("An intentional boundary cannot claim a roadmap reference.", nameof(roadmapReference));
        }
        Operation = operation;
        Disposition = disposition;
        Rationale = rationale.Trim();
        RoadmapReference = normalizedReference;
    }

    /// <summary>Operation whose unsupported state is classified.</summary>
    public OfficeProtectionOperation Operation { get; }

    /// <summary>Whether the operation is roadmap-tracked or deliberately outside the product boundary.</summary>
    public OfficeProtectionUnsupportedDisposition Disposition { get; }

    /// <summary>Concrete explanation of the roadmap outcome or intentional boundary.</summary>
    public string Rationale { get; }

    /// <summary>Roadmap link emitted by the catalog when <see cref="Disposition"/> is roadmap-tracked.</summary>
    public string? RoadmapReference { get; }
}

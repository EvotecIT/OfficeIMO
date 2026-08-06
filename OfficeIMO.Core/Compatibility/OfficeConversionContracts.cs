namespace OfficeIMO;

/// <summary>Controls whether a conversion may continue when content loss is known.</summary>
public enum OfficeConversionLossPolicy {
    /// <summary>Reject conversion when known content would be omitted.</summary>
    Block,

    /// <summary>Continue and report known omitted content in the result.</summary>
    Allow
}

/// <summary>Controls how a conversion handles an existing destination file.</summary>
public enum OfficeConversionFileConflictPolicy {
    /// <summary>Reject conversion if the destination exists.</summary>
    FailIfExists,

    /// <summary>Replace an existing destination through an atomic commit.</summary>
    Replace
}

/// <summary>Identifies the purpose of a conversion diagnostic.</summary>
public enum OfficeConversionDiagnosticCategory {
    /// <summary>Source format detection or extension findings.</summary>
    SourceFormat,

    /// <summary>Content that cannot survive conversion.</summary>
    DataLoss,

    /// <summary>Destination format or writer findings.</summary>
    DestinationFormat
}

/// <summary>Identifies the severity of a conversion diagnostic.</summary>
public enum OfficeConversionDiagnosticSeverity {
    /// <summary>Informational finding.</summary>
    Information,

    /// <summary>Finding requiring user review.</summary>
    Warning,

    /// <summary>Finding that prevented conversion.</summary>
    Error
}

/// <summary>Identifies why a validated document conversion was rejected.</summary>
public enum OfficeConversionFailureReason {
    /// <summary>Source and destination physical formats are identical.</summary>
    SameFormat,

    /// <summary>The destination exists and replacement was not allowed.</summary>
    DestinationExists,

    /// <summary>Known content loss was blocked by policy.</summary>
    DataLossBlocked,

    /// <summary>The destination writer cannot represent source content.</summary>
    DestinationFeatureUnsupported
}

/// <summary>Describes a structured document-conversion finding.</summary>
public sealed class OfficeConversionDiagnostic {
    internal OfficeConversionDiagnostic(
        string code,
        OfficeConversionDiagnosticCategory category,
        OfficeConversionDiagnosticSeverity severity,
        string message,
        bool representsDataLoss,
        OfficeCompatibilityState? compatibilityState = null,
        OfficeCompatibilityImpact compatibilityImpact = OfficeCompatibilityImpact.None,
        string? sourceLocation = null,
        string? fallbackArtifact = null) {
        Code = code;
        Category = category;
        Severity = severity;
        Message = message;
        RepresentsDataLoss = representsDataLoss;
        CompatibilityState = compatibilityState ?? InferCompatibilityState(category, representsDataLoss);
        CompatibilityImpact = compatibilityImpact == OfficeCompatibilityImpact.None && representsDataLoss
            ? OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Carrier
            : compatibilityImpact;
        SourceLocation = sourceLocation;
        FallbackArtifact = fallbackArtifact;
    }

    internal OfficeConversionDiagnostic(
        string code,
        OfficeConversionDiagnosticCategory category,
        OfficeConversionDiagnosticSeverity severity,
        string message,
        OfficeCompatibilityState compatibilityState,
        OfficeCompatibilityImpact compatibilityImpact,
        bool representsDataLoss,
        string? sourceLocation = null,
        string? fallbackArtifact = null)
        : this(
            code,
            category,
            severity,
            message,
            representsDataLoss,
            compatibilityState,
            compatibilityImpact,
            sourceLocation,
            fallbackArtifact) {
    }

    /// <summary>Gets the stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Gets the diagnostic category.</summary>
    public OfficeConversionDiagnosticCategory Category { get; }

    /// <summary>Gets the diagnostic severity.</summary>
    public OfficeConversionDiagnosticSeverity Severity { get; }

    /// <summary>Gets the human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Gets whether the diagnostic describes content that will not survive conversion.</summary>
    public bool RepresentsDataLoss { get; }

    /// <summary>Gets the shared feature-level representation state.</summary>
    public OfficeCompatibilityState CompatibilityState { get; }

    /// <summary>Gets the fidelity dimensions affected by the finding.</summary>
    public OfficeCompatibilityImpact CompatibilityImpact { get; }

    /// <summary>Gets the related source part, story, sheet, range, record, or other location.</summary>
    public string? SourceLocation { get; }

    /// <summary>Gets the generated fallback artifact, when one exists.</summary>
    public string? FallbackArtifact { get; }

    private static OfficeCompatibilityState InferCompatibilityState(
        OfficeConversionDiagnosticCategory category,
        bool representsDataLoss) {
        if (category == OfficeConversionDiagnosticCategory.DestinationFormat) {
            return OfficeCompatibilityState.Blocked;
        }

        return representsDataLoss ? OfficeCompatibilityState.Dropped : OfficeCompatibilityState.Equivalent;
    }
}

/// <summary>Describes how completely OfficeIMO can work with a discovered document feature.</summary>
public enum OfficeFeatureSupportLevel {
    /// <summary>The feature can be read, authored, and edited through the public API.</summary>
    Editable,

    /// <summary>The feature has useful public support but not every authored detail is editable.</summary>
    PartiallyEditable,

    /// <summary>The feature is retained during supported round trips but is not directly editable.</summary>
    Preserved,

    /// <summary>The feature cannot currently be preserved or represented safely.</summary>
    Unsupported
}

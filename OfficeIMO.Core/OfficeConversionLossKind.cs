namespace OfficeIMO;

/// <summary>
/// Classifies the fidelity impact represented by a conversion or export diagnostic.
/// </summary>
public enum OfficeConversionLossKind {
    /// <summary>The diagnostic does not represent fidelity loss.</summary>
    None,

    /// <summary>Content was rendered using a documented approximation.</summary>
    Approximation,

    /// <summary>Source content was omitted or replaced by a fallback.</summary>
    Omission,

    /// <summary>The requested conversion or export operation, or part of it, failed.</summary>
    Failure
}

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes whether an image-export diagnostic represents fidelity loss.
/// </summary>
public enum OfficeImageExportLossKind {
    /// <summary>The diagnostic does not represent fidelity loss.</summary>
    None,

    /// <summary>Content was rendered using a documented approximation.</summary>
    Approximation,

    /// <summary>Source content was omitted or replaced by a fallback.</summary>
    Omission,

    /// <summary>The requested export operation or part of it failed.</summary>
    Failure
}

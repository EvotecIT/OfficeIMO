using System;

namespace OfficeIMO.Drawing;

/// <summary>Describes one feature-level compatibility decision made by a conversion.</summary>
public sealed class OfficeCompatibilityFinding {
    /// <summary>Creates a compatibility finding.</summary>
    public OfficeCompatibilityFinding(
        string code,
        string category,
        string message,
        OfficeCompatibilityState state,
        OfficeCompatibilitySeverity severity = OfficeCompatibilitySeverity.Information,
        OfficeCompatibilityImpact impact = OfficeCompatibilityImpact.None,
        bool representsLoss = false,
        string? sourceLocation = null,
        string? fallbackArtifact = null) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Finding code cannot be empty.", nameof(code));
        if (string.IsNullOrWhiteSpace(category)) throw new ArgumentException("Finding category cannot be empty.", nameof(category));

        Code = code.Trim();
        Category = category.Trim();
        Message = message ?? string.Empty;
        State = state;
        Severity = severity;
        Impact = impact;
        RepresentsLoss = representsLoss;
        SourceLocation = string.IsNullOrWhiteSpace(sourceLocation) ? null : sourceLocation;
        FallbackArtifact = string.IsNullOrWhiteSpace(fallbackArtifact) ? null : fallbackArtifact;
    }

    /// <summary>Gets the stable finding code.</summary>
    public string Code { get; }

    /// <summary>Gets the feature category.</summary>
    public string Category { get; }

    /// <summary>Gets the human-readable compatibility detail.</summary>
    public string Message { get; }

    /// <summary>Gets how the source feature is represented in the target.</summary>
    public OfficeCompatibilityState State { get; }

    /// <summary>Gets the finding severity.</summary>
    public OfficeCompatibilitySeverity Severity { get; }

    /// <summary>Gets the fidelity dimensions affected by the decision.</summary>
    public OfficeCompatibilityImpact Impact { get; }

    /// <summary>Gets whether the decision loses source fidelity under the requested conversion.</summary>
    public bool RepresentsLoss { get; }

    /// <summary>Gets an optional source part, sheet, slide, range, or other location.</summary>
    public string? SourceLocation { get; }

    /// <summary>Gets an optional path or relationship describing a generated fallback artifact.</summary>
    public string? FallbackArtifact { get; }
}

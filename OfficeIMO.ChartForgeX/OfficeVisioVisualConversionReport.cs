using System.Collections.Generic;
using System.Linq;
using global::ChartForgeX.VisualArtifacts;

namespace OfficeIMO.ChartForgeX;

/// <summary>Describes semantic fidelity decisions made during native Visio projection.</summary>
public sealed class OfficeVisioVisualConversionReport {
    private readonly List<OfficeVisioVisualDiagnostic> _diagnostics = new List<OfficeVisioVisualDiagnostic>();

    /// <summary>Gets the broad CFX artifact or authoring kind that was projected.</summary>
    public VisualArtifactKind ArtifactKind { get; internal set; }

    /// <summary>Gets the structured semantic family that selected the native projection.</summary>
    public VisualArtifactInterchangeFamily SemanticFamily { get; internal set; }

    /// <summary>Gets whether every projected object remains independently editable in Visio.</summary>
    public bool AllProjectedObjectsEditable { get; internal set; }

    /// <summary>Gets whether at least one source semantic was not represented exactly.</summary>
    public bool HasSemanticLoss => _diagnostics.Any(item => item.Severity == OfficeVisioVisualDiagnosticSeverity.Warning);

    /// <summary>Gets the OfficeIMO.Visio native builder used for the projection.</summary>
    public OfficeVisioVisualProjectionKind Projection { get; internal set; }

    /// <summary>Gets the number of projected groups or containers.</summary>
    public int GroupCount { get; internal set; }

    /// <summary>Gets the number of projected nodes or participants.</summary>
    public int NodeCount { get; internal set; }

    /// <summary>Gets the number of projected connectors or messages.</summary>
    public int EdgeCount { get; internal set; }

    /// <summary>Gets the number of projected notes or combined fragments.</summary>
    public int AnnotationCount { get; internal set; }

    /// <summary>Gets typed fidelity diagnostics.</summary>
    public IReadOnlyList<OfficeVisioVisualDiagnostic> Diagnostics => _diagnostics;

    /// <summary>Gets human-readable warning messages for logging and interactive display.</summary>
    public IReadOnlyList<string> Warnings => _diagnostics
        .Where(item => item.Severity == OfficeVisioVisualDiagnosticSeverity.Warning)
        .Select(item => item.Message)
        .ToArray();

    internal void Warn(
        OfficeVisioVisualDiagnosticCode code,
        OfficeVisioVisualEntityKind entityKind,
        string? entityId,
        string? feature,
        string message) =>
        _diagnostics.Add(new OfficeVisioVisualDiagnostic(code, OfficeVisioVisualDiagnosticSeverity.Warning, entityKind, entityId, feature, message));

    internal void Info(
        OfficeVisioVisualDiagnosticCode code,
        OfficeVisioVisualEntityKind entityKind,
        string? entityId,
        string? feature,
        string message) =>
        _diagnostics.Add(new OfficeVisioVisualDiagnostic(code, OfficeVisioVisualDiagnosticSeverity.Information, entityKind, entityId, feature, message));

}

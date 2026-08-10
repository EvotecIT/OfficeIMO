using System.Collections.Generic;
using global::ChartForgeX.VisualArtifacts;

namespace OfficeIMO.ChartForgeX;

/// <summary>Describes semantic fidelity decisions made during native Visio projection.</summary>
public sealed class OfficeVisioVisualConversionReport {
    private readonly List<string> _warnings = new List<string>();

    /// <summary>Gets the CFX artifact family that was projected.</summary>
    public VisualArtifactKind ArtifactKind { get; internal set; }

    /// <summary>Gets whether every rendered diagram object remains independently editable in Visio.</summary>
    public bool IsNativeEditable { get; internal set; }

    /// <summary>Gets the OfficeIMO.Visio native builder used for the projection.</summary>
    public string Projection { get; internal set; } = string.Empty;

    /// <summary>Gets the number of projected groups or containers.</summary>
    public int GroupCount { get; internal set; }

    /// <summary>Gets the number of projected nodes or participants.</summary>
    public int NodeCount { get; internal set; }

    /// <summary>Gets the number of projected connectors or messages.</summary>
    public int EdgeCount { get; internal set; }

    /// <summary>Gets the number of projected notes or combined fragments.</summary>
    public int AnnotationCount { get; internal set; }

    /// <summary>Gets human-readable fidelity warnings.</summary>
    public IReadOnlyList<string> Warnings => _warnings;

    internal void Warn(string message) => _warnings.Add(message);
}

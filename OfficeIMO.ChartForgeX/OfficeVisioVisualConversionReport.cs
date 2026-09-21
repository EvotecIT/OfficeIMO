using System;
using System.Collections.Generic;
using System.Linq;
using global::ChartForgeX.VisualArtifacts;

namespace OfficeIMO.ChartForgeX;

/// <summary>Describes semantic fidelity decisions made during native Visio projection.</summary>
public sealed class OfficeVisioVisualConversionReport : IOfficeConversionReport {
    private readonly List<OfficeVisioVisualDiagnostic> _diagnostics = new List<OfficeVisioVisualDiagnostic>();
    private readonly List<OfficeConversionFidelityDiagnostic> _fidelityDiagnostics = new List<OfficeConversionFidelityDiagnostic>();
    private readonly IReadOnlyList<OfficeVisioVisualDiagnostic> _readOnlyDiagnostics;
    private readonly IReadOnlyList<OfficeConversionFidelityDiagnostic> _readOnlyFidelityDiagnostics;

    /// <summary>Creates an empty native Visio projection report.</summary>
    public OfficeVisioVisualConversionReport() {
        _readOnlyDiagnostics = _diagnostics.AsReadOnly();
        _readOnlyFidelityDiagnostics = _fidelityDiagnostics.AsReadOnly();
    }

    /// <summary>Gets the broad CFX artifact or authoring kind that was projected.</summary>
    public VisualArtifactKind ArtifactKind { get; internal set; }

    /// <summary>Gets the structured semantic family that selected the native projection.</summary>
    public VisualArtifactInterchangeFamily SemanticFamily { get; internal set; }

    /// <summary>Gets whether every projected object remains independently editable in Visio.</summary>
    public bool AllProjectedObjectsEditable { get; internal set; }

    /// <summary>Gets whether at least one source semantic was not represented exactly.</summary>
    public bool HasSemanticLoss => HasLoss;

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
    public IReadOnlyList<OfficeVisioVisualDiagnostic> Diagnostics => _readOnlyDiagnostics;

    /// <summary>Gets category-preserving native Visio projection diagnostics.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics => _readOnlyFidelityDiagnostics;

    /// <summary>Gets whether at least one source semantic was approximated or omitted.</summary>
    public bool HasLoss => _fidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>Throws when native Visio projection reported possible fidelity loss.</summary>
    public void RequireNoLoss() {
        if (HasLoss) throw new InvalidOperationException(
            "ChartForgeX native Visio projection reported possible fidelity loss. Inspect FidelityDiagnostics for details.");
    }

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
        AddDiagnostic(code, OfficeVisioVisualDiagnosticSeverity.Warning, entityKind, entityId, feature, message);

    internal void Info(
        OfficeVisioVisualDiagnosticCode code,
        OfficeVisioVisualEntityKind entityKind,
        string? entityId,
        string? feature,
        string message) =>
        AddDiagnostic(code, OfficeVisioVisualDiagnosticSeverity.Information, entityKind, entityId, feature, message);

    private void AddDiagnostic(
        OfficeVisioVisualDiagnosticCode code,
        OfficeVisioVisualDiagnosticSeverity severity,
        OfficeVisioVisualEntityKind entityKind,
        string? entityId,
        string? feature,
        string message) {
        _diagnostics.Add(new OfficeVisioVisualDiagnostic(code, severity, entityKind, entityId, feature, message));
        _fidelityDiagnostics.Add(new OfficeConversionFidelityDiagnostic(
            code.ToString(),
            message,
            ResolveLossKind(code, severity),
            "OfficeIMO.ChartForgeX.Visio",
            ResolveLocation(entityKind, entityId, feature)));
    }

    private static OfficeConversionLossKind ResolveLossKind(
        OfficeVisioVisualDiagnosticCode code,
        OfficeVisioVisualDiagnosticSeverity severity) {
        if (severity == OfficeVisioVisualDiagnosticSeverity.Information) return OfficeConversionLossKind.None;
        return code switch {
            OfficeVisioVisualDiagnosticCode.WatermarkNotProjected or
            OfficeVisioVisualDiagnosticCode.ColorNotProjected or
            OfficeVisioVisualDiagnosticCode.TooltipNotProjected or
            OfficeVisioVisualDiagnosticCode.GroupNotProjected or
            OfficeVisioVisualDiagnosticCode.AnnotationNotProjected or
            OfficeVisioVisualDiagnosticCode.ExtensionsNotProjected or
            OfficeVisioVisualDiagnosticCode.AccessibilityNotProjected or
            OfficeVisioVisualDiagnosticCode.PresentationNotProjected or
            OfficeVisioVisualDiagnosticCode.ScenarioNotProjected or
            OfficeVisioVisualDiagnosticCode.ArtworkNotProjected or
            OfficeVisioVisualDiagnosticCode.EndpointLabelsNotRendered or
            OfficeVisioVisualDiagnosticCode.ShapeDataDisabled or
            OfficeVisioVisualDiagnosticCode.HyperlinkNotProjected or
            OfficeVisioVisualDiagnosticCode.ActivationNotProjected or
            OfficeVisioVisualDiagnosticCode.BranchDividerNotProjected or
            OfficeVisioVisualDiagnosticCode.DetailsNotRendered or
            OfficeVisioVisualDiagnosticCode.TitleNotProjected => OfficeConversionLossKind.Omission,
            _ => OfficeConversionLossKind.Approximation
        };
    }

    private static string ResolveLocation(
        OfficeVisioVisualEntityKind entityKind,
        string? entityId,
        string? feature) {
        string location = entityKind.ToString();
        if (!string.IsNullOrEmpty(entityId)) location += ":" + entityId;
        if (!string.IsNullOrEmpty(feature)) location += "/" + feature;
        return location;
    }

}

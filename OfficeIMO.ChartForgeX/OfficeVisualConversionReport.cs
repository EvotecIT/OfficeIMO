using System;
using System.Collections.Generic;

namespace OfficeIMO.ChartForgeX;

/// <summary>Describes fidelity decisions made while converting one visual artifact.</summary>
public sealed class OfficeVisualConversionReport : OfficeConversionReport {
    private readonly List<string> _warnings = new List<string>();
    private readonly List<OfficeDiagnostic> _fidelityDiagnostics = new List<OfficeDiagnostic>();
    private readonly IReadOnlyList<string> _readOnlyWarnings;
    private readonly IReadOnlyList<OfficeDiagnostic> _readOnlyFidelityDiagnostics;

    /// <summary>Creates an empty visual conversion report.</summary>
    public OfficeVisualConversionReport() {
        _readOnlyWarnings = _warnings.AsReadOnly();
        _readOnlyFidelityDiagnostics = _fidelityDiagnostics.AsReadOnly();
    }

    /// <summary>Gets whether the OfficeDrawing result retains vector content.</summary>
    public bool IsVector { get; internal set; }

    /// <summary>Gets whether the adapter used ChartForgeX PNG output instead of the imported SVG scene.</summary>
    public bool UsedRasterFallback { get; internal set; }

    /// <summary>Gets the number of SVG features the Office drawing importer could not represent completely.</summary>
    public int UnsupportedSvgFeatureCount { get; internal set; }

    /// <summary>Gets human-readable fidelity warnings.</summary>
    public IReadOnlyList<string> Warnings => _readOnlyWarnings;

    /// <summary>Gets category-preserving visual conversion diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics => _readOnlyFidelityDiagnostics;

    /// <summary>Gets whether the conversion approximated or omitted source semantics.</summary>
    public override bool HasLoss => _fidelityDiagnostics.Exists(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>Throws when the conversion reported possible fidelity loss.</summary>
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("ChartForgeX visual conversion reported possible fidelity loss. Inspect FidelityDiagnostics for details.", this);

    internal void Warn(string code, string message, OfficeConversionLossKind lossKind, string? location = null) {
        _warnings.Add(message);
        _fidelityDiagnostics.Add(new global::OfficeIMO.OfficeDiagnostic(
            code,
            lossKind == global::OfficeIMO.OfficeConversionLossKind.None ? global::OfficeIMO.OfficeDiagnosticSeverity.Information : global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
            message,
            location,
            "OfficeIMO.ChartForgeX",
            lossKind));
    }
}

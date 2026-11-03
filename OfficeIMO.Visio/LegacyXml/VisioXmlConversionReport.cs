using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio;

/// <summary>Fidelity diagnostics for one legacy Visio XML import or export.</summary>
public sealed class VisioXmlConversionReport : OfficeConversionReport {
    private readonly List<OfficeDiagnostic> _diagnostics = new();
    /// <inheritdoc />
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics => _diagnostics.AsReadOnly();
    /// <inheritdoc />
    public override bool HasLoss => _diagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
    /// <summary>Whether the operation omitted source content.</summary>
    public bool HasOmissions => _diagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Omission);
    /// <inheritdoc />
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("Legacy Visio XML conversion reported fidelity loss.", this);
    internal void Add(string code, string message, OfficeConversionLossKind kind = OfficeConversionLossKind.Omission, string? location = null) {
        if (!_diagnostics.Any(item => item.Code == code && item.Location == location && item.Message == message))
            _diagnostics.Add(new global::OfficeIMO.OfficeDiagnostic(code, kind == global::OfficeIMO.OfficeConversionLossKind.None ? global::OfficeIMO.OfficeDiagnosticSeverity.Information : global::OfficeIMO.OfficeDiagnosticSeverity.Warning, message, location, "OfficeIMO.Visio.LegacyXml", kind));
    }
}

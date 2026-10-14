using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio;

/// <summary>Fidelity diagnostics for one legacy Visio XML import or export.</summary>
public sealed class VisioXmlConversionReport : IOfficeConversionReport {
    private readonly List<OfficeConversionFidelityDiagnostic> _diagnostics = new();
    /// <inheritdoc />
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics => _diagnostics.AsReadOnly();
    /// <inheritdoc />
    public bool HasLoss => _diagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
    /// <summary>Whether the operation omitted source content.</summary>
    public bool HasOmissions => _diagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Omission);
    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new OfficeConversionException("Legacy Visio XML conversion reported fidelity loss.", this);
    }
    internal void Add(string code, string message, OfficeConversionLossKind kind = OfficeConversionLossKind.Omission, string? location = null) {
        if (!_diagnostics.Any(item => item.Code == code && item.Location == location && item.Message == message))
            _diagnostics.Add(new OfficeConversionFidelityDiagnostic(code, message, kind, "OfficeIMO.Visio.LegacyXml", location));
    }
}

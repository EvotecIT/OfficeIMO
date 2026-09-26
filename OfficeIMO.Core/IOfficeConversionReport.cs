using System.Collections.Generic;

namespace OfficeIMO;

/// <summary>
/// Common fidelity contract for reports produced by one stage of an Office document conversion.
/// </summary>
public interface IOfficeConversionReport {
    /// <summary>
    /// Gets immutable, category-preserving diagnostics for this conversion stage.
    /// </summary>
    IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }

    /// <summary>True when the conversion stage approximated, omitted, or could not preserve source content.</summary>
    bool HasLoss { get; }

    /// <summary>Throws when the conversion stage reported possible content loss.</summary>
    void RequireNoLoss();
}

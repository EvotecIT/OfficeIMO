using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>
/// Acceptance policy applied to structured image-export diagnostics.
/// </summary>
public sealed class OfficeImageExportPolicy {
    private IReadOnlyCollection<string> _failOnDiagnosticCodes = Array.Empty<string>();

    /// <summary>Rejects any approximation, omission, or failure diagnostic.</summary>
    public bool RequireNoLoss { get; set; }

    /// <summary>Rejects omission diagnostics while allowing documented approximations.</summary>
    public bool RequireNoOmissions { get; set; }

    /// <summary>Rejects error/failure diagnostics.</summary>
    public bool RequireNoFailures { get; set; } = true;

    /// <summary>Stable diagnostic codes that should fail the export.</summary>
    public IReadOnlyCollection<string> FailOnDiagnosticCodes {
        get => _failOnDiagnosticCodes;
        set => _failOnDiagnosticCodes = NormalizeCodes(value);
    }

    /// <summary>Creates an independent policy snapshot.</summary>
    public OfficeImageExportPolicy Clone() => new OfficeImageExportPolicy {
        RequireNoLoss = RequireNoLoss,
        RequireNoOmissions = RequireNoOmissions,
        RequireNoFailures = RequireNoFailures,
        FailOnDiagnosticCodes = _failOnDiagnosticCodes
    };

    /// <summary>Throws when the diagnostics violate this policy.</summary>
    public void EnsureAccepted(IEnumerable<OfficeImageExportDiagnostic> diagnostics) {
        if (diagnostics == null) throw new ArgumentNullException(nameof(diagnostics));

        OfficeImageExportDiagnostic[] rejected = diagnostics.Where(IsRejected).ToArray();
        if (rejected.Length == 0) return;
        throw new OfficeImageExportPolicyException(rejected);
    }

    private bool IsRejected(OfficeImageExportDiagnostic diagnostic) {
        if (diagnostic == null) return false;
        if (_failOnDiagnosticCodes.Contains(diagnostic.Code, StringComparer.OrdinalIgnoreCase)) return true;
        if (RequireNoLoss && diagnostic.LossKind != OfficeImageExportLossKind.None) return true;
        if (RequireNoOmissions && diagnostic.LossKind == OfficeImageExportLossKind.Omission) return true;
        return RequireNoFailures &&
               (diagnostic.LossKind == OfficeImageExportLossKind.Failure ||
                diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
    }

    private static IReadOnlyCollection<string> NormalizeCodes(IEnumerable<string>? codes) {
        if (codes == null) return Array.Empty<string>();
        return codes
            .Where(code => !string.IsNullOrWhiteSpace(code))
            .Select(code => code.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }
}

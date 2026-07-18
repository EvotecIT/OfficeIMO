using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>
/// Raised when structured image-export diagnostics violate the caller's acceptance policy.
/// </summary>
public sealed class OfficeImageExportPolicyException : InvalidOperationException {
    internal OfficeImageExportPolicyException(IReadOnlyList<OfficeImageExportDiagnostic> diagnostics)
        : base(CreateMessage(diagnostics)) {
        Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
    }

    /// <summary>Diagnostics rejected by the policy.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }

    private static string CreateMessage(IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
        if (diagnostics == null || diagnostics.Count == 0) return "Image export did not satisfy the configured acceptance policy.";
        return "Image export did not satisfy the configured acceptance policy: " +
               string.Join(", ", diagnostics.Select(diagnostic => diagnostic.Code)) + ".";
    }
}

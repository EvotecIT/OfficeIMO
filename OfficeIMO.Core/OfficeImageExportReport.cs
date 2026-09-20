using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>
/// Aggregate diagnostics and fidelity status for one or more image-export results.
/// </summary>
public sealed class OfficeImageExportReport : global::OfficeIMO.IOfficeConversionReport {
    private readonly ReadOnlyCollection<OfficeImageExportDiagnostic> _diagnostics;

    /// <summary>Creates an aggregate report from result diagnostics.</summary>
    public OfficeImageExportReport(IEnumerable<OfficeImageExportResult> results) {
        if (results == null) throw new ArgumentNullException(nameof(results));
        OfficeImageExportResult[] snapshot = results.ToArray();
        ResultCount = snapshot.Length;
        _diagnostics = Array.AsReadOnly(snapshot.SelectMany(result => result.Diagnostics).ToArray());
    }

    internal OfficeImageExportReport(IEnumerable<OfficeImageExportDiagnostic> diagnostics, int resultCount) {
        if (diagnostics == null) throw new ArgumentNullException(nameof(diagnostics));
        if (resultCount < 0) throw new ArgumentOutOfRangeException(nameof(resultCount));
        ResultCount = resultCount;
        _diagnostics = Array.AsReadOnly(diagnostics.ToArray());
    }

    /// <summary>Number of exported results represented by the report.</summary>
    public int ResultCount { get; }

    /// <summary>All structured diagnostics in result order.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics => _diagnostics;

    /// <summary>Diagnostics projected onto the common typed conversion contract.</summary>
    public IReadOnlyList<global::OfficeIMO.OfficeConversionFidelityDiagnostic> FidelityDiagnostics =>
        Array.AsReadOnly(_diagnostics.Select(diagnostic => new global::OfficeIMO.OfficeConversionFidelityDiagnostic(
            diagnostic.Code,
            diagnostic.Message,
            diagnostic.LossKind,
            diagnostic.FidelitySource ?? "OfficeIMO.Drawing",
            diagnostic.FidelityLocation ?? diagnostic.Source)).ToArray());

    /// <summary>True when at least one diagnostic represents fidelity loss.</summary>
    public bool HasLoss => _diagnostics.Any(diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>True when at least one source feature was omitted or replaced.</summary>
    public bool HasOmissions => _diagnostics.Any(diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);

    /// <summary>True when at least one requested operation or part failed.</summary>
    public bool HasFailures => _diagnostics.Any(diagnostic =>
        diagnostic.LossKind == OfficeConversionLossKind.Failure ||
        diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);

    /// <summary>Applies an acceptance policy to this report.</summary>
    public OfficeImageExportReport Require(OfficeImageExportPolicy policy) {
        if (policy == null) throw new ArgumentNullException(nameof(policy));
        policy.EnsureAccepted(_diagnostics);
        return this;
    }

    /// <summary>Throws when any export diagnostic represents fidelity loss.</summary>
    public void RequireNoLoss() {
        if (HasLoss) throw new InvalidOperationException("Image export reported fidelity loss. Inspect FidelityDiagnostics.");
    }
}

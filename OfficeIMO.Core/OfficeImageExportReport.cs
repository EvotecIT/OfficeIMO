using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>
/// Aggregate diagnostics and fidelity status for one or more image-export results.
/// </summary>
public sealed class OfficeImageExportReport : OfficeConversionReport {
    private readonly ReadOnlyCollection<OfficeImageExportDiagnostic> _diagnostics;

    /// <summary>Creates an aggregate report from result diagnostics.</summary>
    public OfficeImageExportReport(IEnumerable<OfficeImageExportResult> results) {
        ArgumentNullException.ThrowIfNull(results);
        OfficeImageExportResult[] snapshot = results.ToArray();
        ResultCount = snapshot.Length;
        _diagnostics = Array.AsReadOnly(snapshot.SelectMany(result => result.Diagnostics).ToArray());
    }

    internal OfficeImageExportReport(IEnumerable<OfficeImageExportDiagnostic> diagnostics, int resultCount) {
        ArgumentNullException.ThrowIfNull(diagnostics);
        ArgumentOutOfRangeException.ThrowIfNegative(resultCount);
        ResultCount = resultCount;
        _diagnostics = Array.AsReadOnly(diagnostics.ToArray());
    }

    /// <summary>Number of exported results represented by the report.</summary>
    public int ResultCount { get; }

    /// <summary>All structured diagnostics in result order.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> NativeDiagnostics => _diagnostics;

    /// <summary>NativeDiagnostics projected onto the common typed conversion contract.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics =>
        Array.AsReadOnly(_diagnostics.Select(diagnostic => new global::OfficeIMO.OfficeDiagnostic(
            diagnostic.Code,
            diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.None ? global::OfficeIMO.OfficeDiagnosticSeverity.Information : global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
            diagnostic.Message,
            diagnostic.FidelityLocation ?? diagnostic.Source,
            diagnostic.FidelitySource ?? "OfficeIMO.Drawing",
            diagnostic.LossKind)).ToArray());

    /// <summary>True when at least one diagnostic represents fidelity loss.</summary>
    public override bool HasLoss => _diagnostics.Any(diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>True when at least one source feature was omitted or replaced.</summary>
    public bool HasOmissions => _diagnostics.Any(diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);

    /// <summary>True when at least one requested operation or part failed.</summary>
    public bool HasFailures => _diagnostics.Any(diagnostic =>
        diagnostic.LossKind == OfficeConversionLossKind.Failure ||
        diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);

    /// <summary>Applies an acceptance policy to this report.</summary>
    public OfficeImageExportReport Require(OfficeImageExportPolicy policy) {
        ArgumentNullException.ThrowIfNull(policy);
        policy.EnsureAccepted(_diagnostics);
        return this;
    }

    /// <summary>Throws when any export diagnostic represents fidelity loss.</summary>
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("Image export reported fidelity loss. Inspect Diagnostics.", this);
}

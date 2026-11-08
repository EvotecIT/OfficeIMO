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
    public OfficeImageExportReport(IEnumerable<OfficeImageExportResult> results) : this(Snapshot(results)) { }

    private static OfficeImageExportResult[] Snapshot(IEnumerable<OfficeImageExportResult> results) {
        ArgumentNullException.ThrowIfNull(results);
        OfficeImageExportResult[] snapshot = results.ToArray();
        if (snapshot.Any(result => result == null)) throw new ArgumentException("Results cannot contain null entries.", nameof(results));
        return snapshot;
    }

    private OfficeImageExportReport(OfficeImageExportResult[] results) : this(
        results.SelectMany(result => result.Diagnostics).ToArray(), results.Length,
        results.Select(result => result.Source ?? "Image").Distinct().Count() == 1 ? results[0].Source ?? "Image" : "Image",
        results.Select(result => result.Format).Distinct().Count() == 1 ? results[0].Format.ToString() : "Images",
        results.Length == 1 ? results[0].SavedPath : null) { }

    internal OfficeImageExportReport(IEnumerable<OfficeImageExportDiagnostic> diagnostics, int resultCount) : this(
        (diagnostics ?? throw new ArgumentNullException(nameof(diagnostics))).ToArray(), resultCount, "Image", "Images", null) { }

    private OfficeImageExportReport(OfficeImageExportDiagnostic[] diagnostics, int resultCount, string source, string target, string? outputPath)
        : base(source, target, outputPath, resultCount, diagnostics: diagnostics.Select(diagnostic => OfficeImageExportEvidence.ProjectDiagnostic(source, diagnostic))) {
        ResultCount = resultCount;
        _diagnostics = Array.AsReadOnly(diagnostics);
    }

    /// <summary>Number of exported results represented by the report.</summary>
    public int ResultCount { get; }

    /// <summary>All structured diagnostics in result order.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> NativeDiagnostics => _diagnostics;

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

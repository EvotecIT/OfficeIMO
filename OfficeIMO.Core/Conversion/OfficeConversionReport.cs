using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO;

/// <summary>An immutable record of a completed conversion. Conversion failures throw rather than return a failed report.</summary>
public sealed class OfficeConversionReport : IOfficeConversionReport {
    /// <summary>Creates a report, taking a defensive snapshot of all diagnostics.</summary>
    /// <param name="source">Source path, format, or in-memory source description.</param>
    /// <param name="target">Target path, format, or in-memory target description.</param>
    /// <param name="outputPath">Written file path, or null for an in-memory or stream conversion.</param>
    /// <param name="itemsWritten">Number of pages or logical items written, as defined by the converter.</param>
    /// <param name="elapsed">Elapsed conversion time.</param>
    /// <param name="diagnostics">Completed-stage diagnostics, or null for none.</param>
    public OfficeConversionReport(string source, string target, string? outputPath = null, long itemsWritten = 0,
        TimeSpan elapsed = default, IEnumerable<OfficeDiagnostic>? diagnostics = null) {
        if (string.IsNullOrWhiteSpace(source)) throw new ArgumentException("A source is required.", nameof(source));
        if (string.IsNullOrWhiteSpace(target)) throw new ArgumentException("A target is required.", nameof(target));
        if (itemsWritten < 0) throw new ArgumentOutOfRangeException(nameof(itemsWritten));
        if (elapsed < TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(elapsed));
        var snapshot = diagnostics == null ? new List<OfficeDiagnostic>() : new List<OfficeDiagnostic>(diagnostics);
        if (snapshot.Any(d => d == null)) throw new ArgumentException("Diagnostics cannot contain null entries.", nameof(diagnostics));
        Source = source;
        Target = target;
        OutputPath = outputPath;
        ItemsWritten = itemsWritten;
        Elapsed = elapsed;
        Diagnostics = snapshot.AsReadOnly();
        HasLoss = snapshot.Any(d => d.LossKind != OfficeConversionLossKind.None);
    }
    /// <summary>Gets the source path, format or description.</summary>
    public string Source { get; }
    /// <summary>Gets the target path, format or description.</summary>
    public string Target { get; }
    /// <summary>Gets the written path, or null for a stream or in-memory result.</summary>
    public string? OutputPath { get; }
    /// <summary>Gets pages or logical items written, as defined by the converter.</summary>
    public long ItemsWritten { get; }
    /// <summary>Gets elapsed conversion time.</summary>
    public TimeSpan Elapsed { get; }
    /// <summary>Gets the immutable diagnostic snapshot.</summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics { get; }
    /// <inheritdoc />
    public bool HasLoss { get; }
    /// <summary>Throws an <see cref="OfficeConversionException"/> carrying this report if any fidelity loss is reported.</summary>
    public void ThrowIfLoss() {
        if (HasLoss) throw new OfficeConversionException("The conversion reported fidelity loss. Inspect the report diagnostics for details.", this);
    }

    // Explicit legacy adapters keep existing result consumers compiling until migration.
    void IOfficeConversionReport.RequireNoLoss() => ThrowIfLoss();
    IReadOnlyList<OfficeConversionFidelityDiagnostic> IOfficeConversionReport.FidelityDiagnostics =>
        Array.AsReadOnly(Diagnostics.Select(d => new OfficeConversionFidelityDiagnostic(d.Code, d.Message,
            d.LossKind, d.Format ?? Source, d.Location)).ToArray());
}

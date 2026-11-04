using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO;

/// <summary>
/// The common record of a completed conversion or format stage. Format-owned reports derive from it and add their typed evidence.
/// Conversion failures throw rather than return a failed report.
/// </summary>
public class OfficeConversionReport {
    private readonly IReadOnlyList<OfficeDiagnostic>? _diagnostics;

    /// <summary>Creates a report, taking a defensive snapshot of all diagnostics.</summary>
    /// <param name="source">Source path, format, or in-memory source description.</param>
    /// <param name="target">Target path, format, or in-memory target description.</param>
    /// <param name="outputPath">Written file path, or null for an in-memory or stream conversion.</param>
    /// <param name="itemsWritten">Number of pages or logical items written, as defined by the converter.</param>
    /// <param name="elapsed">Elapsed conversion time.</param>
    /// <param name="diagnostics">Completed-stage diagnostics, or null for none.</param>
    public OfficeConversionReport(string source, string target, string? outputPath = null, long itemsWritten = 0,
        TimeSpan elapsed = default, IEnumerable<OfficeDiagnostic>? diagnostics = null) : this(source, target) {
        if (itemsWritten < 0) throw new ArgumentOutOfRangeException(nameof(itemsWritten));
        if (elapsed < TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(elapsed));
        var snapshot = diagnostics == null ? new List<OfficeDiagnostic>() : new List<OfficeDiagnostic>(diagnostics);
        if (snapshot.Any(d => d == null)) throw new ArgumentException("Diagnostics cannot contain null entries.", nameof(diagnostics));
        OutputPath = outputPath;
        ItemsWritten = itemsWritten;
        Elapsed = elapsed;
        _diagnostics = snapshot.AsReadOnly();
    }

    /// <summary>Creates a format-owned report whose diagnostics come from <see cref="GetDiagnostics"/>.</summary>
    /// <param name="source">Source format or description.</param>
    /// <param name="target">Target format or description.</param>
    protected OfficeConversionReport(string source, string target) {
        if (string.IsNullOrWhiteSpace(source)) throw new ArgumentException("A source is required.", nameof(source));
        if (string.IsNullOrWhiteSpace(target)) throw new ArgumentException("A target is required.", nameof(target));
        Source = source;
        Target = target;
    }

    /// <summary>Creates a format-owned report named after the format that owns the derived type.</summary>
    protected OfficeConversionReport() {
        string format = FormatOf(GetType());
        Source = format;
        Target = format;
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
    /// <summary>Gets the common diagnostics. Format-owned reports keep their typed findings in their own members.</summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics => _diagnostics ?? GetDiagnostics()
        ?? throw new InvalidOperationException($"{GetType().FullName} returned null diagnostics.");
    /// <summary>Gets whether the report describes an approximation, omission, failure or unassessed content.</summary>
    public virtual bool HasLoss => Diagnostics.Any(static d => d.LossKind != OfficeConversionLossKind.None);

    /// <summary>Throws an <see cref="OfficeConversionException"/> carrying this report if any fidelity loss is reported.</summary>
    public void ThrowIfLoss() {
        if (HasLoss) throw CreateLossException();
    }

    /// <summary>Produces the common diagnostics of a format-owned report.</summary>
    protected virtual IReadOnlyList<OfficeDiagnostic> GetDiagnostics() => Array.Empty<OfficeDiagnostic>();

    /// <summary>Creates the exception thrown by <see cref="ThrowIfLoss"/>. Format-owned reports can attach typed evidence.</summary>
    protected virtual OfficeConversionException CreateLossException() =>
        new OfficeConversionException("The conversion reported fidelity loss. Inspect the report diagnostics for details.", this);

    private static string FormatOf(Type type) {
        string? ns = type.Namespace;
        if (string.IsNullOrEmpty(ns) || ns == "OfficeIMO") return "OfficeIMO";
        return ns!.StartsWith("OfficeIMO.", StringComparison.Ordinal) ? ns.Substring("OfficeIMO.".Length) : ns;
    }
}

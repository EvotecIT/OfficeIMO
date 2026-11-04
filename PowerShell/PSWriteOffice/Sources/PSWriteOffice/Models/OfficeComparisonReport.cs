using System.Collections.Generic;
using System.Linq;
using OfficeIMO;

namespace PSWriteOffice.Models;

/// <summary>Common comparison summary retaining the complete native findings and conversion evidence.</summary>
public sealed class OfficeComparisonReport {
    internal OfficeComparisonReport(object nativeReport, bool areEqual, IEnumerable<object> findings, int? maximum,
        IReadOnlyList<OfficeConversionReport>? conversions = null) {
        NativeReport = nativeReport; AreEqual = areEqual;
        var all = findings.ToArray(); TotalDifferences = all.Length;
        Differences = maximum.HasValue ? all.Take(maximum.Value).ToArray() : all;
        ConversionReports = conversions ?? System.Array.Empty<OfficeConversionReport>();
    }
    /// <summary>Complete native comparison report, including format-specific diagnostics and artifacts.</summary>
    public object NativeReport { get; }
    /// <summary>Native comparison result; incomplete renderings never qualify as equal.</summary>
    public bool AreEqual { get; }
    /// <summary>Returned findings, optionally limited by MaximumDifferences.</summary>
    public IReadOnlyList<object> Differences { get; }
    /// <summary>Native findings observed before the output limit.</summary>
    public int TotalDifferences { get; }
    /// <summary>Whether the returned finding list was truncated.</summary>
    public bool Truncated => Differences.Count < TotalDifferences;
    /// <summary>Evidence for native conversions performed before a visual comparison.</summary>
    public IReadOnlyList<OfficeConversionReport> ConversionReports { get; }
    /// <summary>Comparison artifact actually published; null includes WhatIf refusal.</summary>
    public string? OutputPath { get; internal set; }
}

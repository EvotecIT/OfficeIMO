using System.Collections.ObjectModel;

namespace OfficeIMO.Pdf;

/// <summary>
/// Stable summary of warnings captured during one PDF conversion workflow.
/// </summary>
public sealed class PdfConversionReportSummary {
    internal PdfConversionReportSummary(IReadOnlyList<PdfConversionWarning> warnings) {
        TotalCount = warnings.Count;

        var severityCounts = new Dictionary<PdfConversionWarningSeverity, int>();
        var converterCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var codeCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var sourceCounts = new Dictionary<string, int>(StringComparer.Ordinal);

        for (int i = 0; i < warnings.Count; i++) {
            PdfConversionWarning warning = warnings[i];
            AddCount(severityCounts, warning.Severity);
            AddCount(converterCounts, NormalizeKey(warning.Converter));
            AddCount(codeCounts, NormalizeKey(warning.Code));

            if (!string.IsNullOrWhiteSpace(warning.Source)) {
                AddCount(sourceCounts, warning.Source);
            }
        }

        InformationCount = GetCount(severityCounts, PdfConversionWarningSeverity.Information);
        WarningCount = GetCount(severityCounts, PdfConversionWarningSeverity.Warning);
        ErrorCount = GetCount(severityCounts, PdfConversionWarningSeverity.Error);
        SeverityCounts = new ReadOnlyDictionary<PdfConversionWarningSeverity, int>(severityCounts);
        ConverterCounts = new ReadOnlyDictionary<string, int>(converterCounts);
        CodeCounts = new ReadOnlyDictionary<string, int>(codeCounts);
        SourceCounts = new ReadOnlyDictionary<string, int>(sourceCounts);
    }

    /// <summary>Total number of warning records in the report.</summary>
    public int TotalCount { get; }

    /// <summary>Number of informational records in the report.</summary>
    public int InformationCount { get; }

    /// <summary>Number of warning-level records in the report.</summary>
    public int WarningCount { get; }

    /// <summary>Number of error-level records in the report.</summary>
    public int ErrorCount { get; }

    /// <summary>True when the report contains any warning, information, or error records.</summary>
    public bool HasWarnings => TotalCount > 0;

    /// <summary>True when the report contains at least one error-level conversion warning.</summary>
    public bool HasErrors => ErrorCount > 0;

    /// <summary>Counts grouped by warning severity.</summary>
    public IReadOnlyDictionary<PdfConversionWarningSeverity, int> SeverityCounts { get; }

    /// <summary>Counts grouped by converter or adapter name.</summary>
    public IReadOnlyDictionary<string, int> ConverterCounts { get; }

    /// <summary>Counts grouped by stable warning code.</summary>
    public IReadOnlyDictionary<string, int> CodeCounts { get; }

    /// <summary>Counts grouped by source area when the warning declared one.</summary>
    public IReadOnlyDictionary<string, int> SourceCounts { get; }

    private static void AddCount<TKey>(Dictionary<TKey, int> counts, TKey key) where TKey : notnull {
        counts[key] = counts.TryGetValue(key, out int count) ? count + 1 : 1;
    }

    private static int GetCount(Dictionary<PdfConversionWarningSeverity, int> counts, PdfConversionWarningSeverity severity) {
        return counts.TryGetValue(severity, out int count) ? count : 0;
    }

    private static string NormalizeKey(string value) {
        return string.IsNullOrWhiteSpace(value) ? "Unknown" : value;
    }
}

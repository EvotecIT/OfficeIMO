namespace PSWriteOffice.Models;

/// <summary>Hidden data found in one file, what removal did with each item, and where the cleaned copy was written.</summary>
public sealed class OfficeHiddenDataFileReport {
    internal OfficeHiddenDataFileReport(OfficeHiddenDataReport report, string path, string? outputPath) {
        Path = path;
        OutputPath = outputPath;
        Format = report.Format;
        Findings = report.Findings;
        Kinds = report.Kinds;
        Changed = report.Changed;
    }

    /// <summary>The inspected file.</summary>
    public string Path { get; }
    /// <summary>The cleaned copy, or null when the file was only inspected.</summary>
    public string? OutputPath { get; }
    /// <summary>The inspected format, such as <c>Word</c>, <c>Excel</c> or <c>PDF</c>.</summary>
    public string Format { get; }
    /// <summary>Every finding in document order, with its outcome.</summary>
    public IReadOnlyList<OfficeHiddenDataFinding> Findings { get; }
    /// <summary>Categories that have at least one finding.</summary>
    public OfficeHiddenDataKind Kinds { get; }
    /// <summary>Whether removal changed anything; always false for an inspection.</summary>
    public bool Changed { get; }
}

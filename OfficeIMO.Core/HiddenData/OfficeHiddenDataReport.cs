using System.Collections.ObjectModel;
using OfficeIMO.Provenance;

namespace OfficeIMO;

/// <summary>One piece of metadata or hidden data found in a document, and what happened to it.</summary>
public sealed class OfficeHiddenDataFinding {
    internal const int MaxValueLength = 160;

    internal OfficeHiddenDataFinding(OfficeHiddenDataKind kind, string location, string description, string? value,
        OfficeHiddenDataOutcome outcome, string? note, OfficeProvenanceEvidence? provenance = null) {
        Kind = kind;
        Location = location;
        Description = description;
        Value = Truncate(value);
        Outcome = outcome;
        Note = note;
        Provenance = provenance;
    }

    /// <summary>The category.</summary>
    public OfficeHiddenDataKind Kind { get; }
    /// <summary>Where it was found: a package part, sheet cell, slide or document story.</summary>
    public string Location { get; }
    /// <summary>What was found.</summary>
    public string Description { get; }
    /// <summary>A short preview of the value, such as an author name or path, when there is one.</summary>
    public string? Value { get; }
    /// <summary>What happened to it.</summary>
    public OfficeHiddenDataOutcome Outcome { get; }
    /// <summary>Why a finding was retained, or other detail about the outcome.</summary>
    public string? Note { get; }
    /// <summary>The provenance carrier behind a <see cref="OfficeHiddenDataKind.Provenance"/> finding, as the provenance inspector reports it.</summary>
    public OfficeProvenanceEvidence? Provenance { get; }

    /// <inheritdoc />
    public override string ToString() => Kind + " " + Location + ": " + Description + " (" + Outcome + ")";

    private static string? Truncate(string? value) {
        if (string.IsNullOrEmpty(value)) return null;
        return value!.Length <= MaxValueLength ? value : value.Substring(0, MaxValueLength) + "…";
    }
}

/// <summary>Metadata, hidden data and provenance found in a Word, Excel or PowerPoint document, and what removal did with each item.</summary>
public sealed class OfficeHiddenDataReport {
    internal OfficeHiddenDataReport(string format, IList<OfficeHiddenDataFinding> findings) {
        Format = format;
        Findings = new ReadOnlyCollection<OfficeHiddenDataFinding>(findings);
    }

    /// <summary>The inspected format, such as <c>Word</c>.</summary>
    public string Format { get; }

    /// <summary>Findings in document order, grouped by category.</summary>
    public IReadOnlyList<OfficeHiddenDataFinding> Findings { get; }

    /// <summary>Whether any metadata or hidden data was found.</summary>
    public bool HasHiddenData => Findings.Count > 0;

    /// <summary>Categories that have at least one finding.</summary>
    public OfficeHiddenDataKind Kinds {
        get {
            OfficeHiddenDataKind kinds = OfficeHiddenDataKind.None;
            foreach (OfficeHiddenDataFinding finding in Findings) kinds |= finding.Kind;
            return kinds;
        }
    }

    /// <summary>Whether removal changed the document.</summary>
    public bool Changed => Findings.Any(finding => finding.Outcome is not (OfficeHiddenDataOutcome.Found or OfficeHiddenDataOutcome.Retained));

    /// <summary>Findings in one category.</summary>
    public IReadOnlyList<OfficeHiddenDataFinding> GetFindings(OfficeHiddenDataKind kind) =>
        Findings.Where(finding => finding.Kind == kind).ToList();
}

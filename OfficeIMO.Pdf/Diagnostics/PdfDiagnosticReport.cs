namespace OfficeIMO.Pdf;

/// <summary>Combined PDF inspection, parser, stream, and preflight diagnostics.</summary>
public sealed class PdfDiagnosticReport {
    internal PdfDiagnosticReport(
        PdfDocumentProbe probe,
        PdfDocumentPreflight preflight,
        PdfDocumentInfo? info,
        IReadOnlyDictionary<string, int> objectTypeCounts,
        IReadOnlyDictionary<string, int> streamTypeCounts,
        IReadOnlyList<PdfStreamDiagnostic> streams,
        IReadOnlyList<PdfDiagnosticFinding> findings,
        bool objectGraphParsed,
        string? objectGraphError) {
        Probe = probe;
        Preflight = preflight;
        Info = info;
        ObjectTypeCounts = objectTypeCounts;
        StreamTypeCounts = streamTypeCounts;
        Streams = streams;
        Findings = findings;
        ObjectGraphParsed = objectGraphParsed;
        ObjectGraphError = objectGraphError;
    }

    /// <summary>Lightweight marker probe.</summary>
    public PdfDocumentProbe Probe { get; }

    /// <summary>Read and rewrite preflight result.</summary>
    public PdfDocumentPreflight Preflight { get; }

    /// <summary>Full readable document info, when parsing succeeded.</summary>
    public PdfDocumentInfo? Info { get; }

    /// <summary>True when OfficeIMO.Pdf can read this document.</summary>
    public bool CanRead => Preflight.CanRead;

    /// <summary>True when OfficeIMO.Pdf can rewrite this document without known blockers.</summary>
    public bool CanRewrite => Preflight.CanRewrite;

    /// <summary>PDF header version, when present.</summary>
    public string? HeaderVersion => Probe.HeaderVersion;

    /// <summary>Page count when known.</summary>
    public int? PageCount => Info?.PageCount;

    /// <summary>True when encryption markers were detected.</summary>
    public bool HasEncryption => Probe.HasEncryption;

    /// <summary>True when signatures were detected.</summary>
    public bool HasSignatures => Probe.HasSignatures;

    /// <summary>True when forms were detected.</summary>
    public bool HasForms => Probe.HasForms;

    /// <summary>True when annotations were detected.</summary>
    public bool HasAnnotations => Probe.HasAnnotations;

    /// <summary>True when outlines/bookmarks were detected.</summary>
    public bool HasOutlines => Probe.HasOutlines;

    /// <summary>True when catalog page labels were detected.</summary>
    public bool HasPageLabels => Probe.HasPageLabels;

    /// <summary>True when catalog open actions were detected.</summary>
    public bool HasOpenActions => Probe.HasOpenActions;

    /// <summary>True when catalog viewer preferences were detected.</summary>
    public bool HasViewerPreferences => Probe.HasViewerPreferences;

    /// <summary>Count of parsed indirect objects by object kind.</summary>
    public IReadOnlyDictionary<string, int> ObjectTypeCounts { get; }

    /// <summary>Count of parsed stream objects by stream kind.</summary>
    public IReadOnlyDictionary<string, int> StreamTypeCounts { get; }

    /// <summary>Parsed stream summaries.</summary>
    public IReadOnlyList<PdfStreamDiagnostic> Streams { get; }

    /// <summary>Total parsed indirect object count.</summary>
    public int ObjectCount {
        get {
            int count = 0;
            foreach (var item in ObjectTypeCounts) {
                count += item.Value;
            }

            return count;
        }
    }

    /// <summary>Total parsed stream count.</summary>
    public int StreamCount => Streams.Count;

    /// <summary>Diagnostic findings.</summary>
    public IReadOnlyList<PdfDiagnosticFinding> Findings { get; }

    /// <summary>True when the indirect object graph was parsed.</summary>
    public bool ObjectGraphParsed { get; }

    /// <summary>Object parser error, when parsing was blocked or unsupported.</summary>
    public string? ObjectGraphError { get; }
}

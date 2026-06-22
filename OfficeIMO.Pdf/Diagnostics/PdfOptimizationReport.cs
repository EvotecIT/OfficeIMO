namespace OfficeIMO.Pdf;

/// <summary>Optimization opportunities discovered without modifying the PDF.</summary>
public sealed class PdfOptimizationReport {
    internal PdfOptimizationReport(
        PdfDiagnosticReport diagnostics,
        IReadOnlyList<PdfDuplicateStreamGroup> duplicateStreams,
        IReadOnlyList<PdfStreamDiagnostic> largestStreams,
        IReadOnlyList<PdfDiagnosticFinding> findings,
        long estimatedSavingsBytes) {
        Diagnostics = diagnostics;
        DuplicateStreams = duplicateStreams;
        LargestStreams = largestStreams;
        Findings = findings;
        EstimatedSavingsBytes = estimatedSavingsBytes;
    }

    /// <summary>Underlying diagnostic report.</summary>
    public PdfDiagnosticReport Diagnostics { get; }

    /// <summary>Duplicate stream candidate groups.</summary>
    public IReadOnlyList<PdfDuplicateStreamGroup> DuplicateStreams { get; }

    /// <summary>Largest stream objects by retained byte length.</summary>
    public IReadOnlyList<PdfStreamDiagnostic> LargestStreams { get; }

    /// <summary>Optimization findings and hints.</summary>
    public IReadOnlyList<PdfDiagnosticFinding> Findings { get; }

    /// <summary>Conservative estimate of bytes that might be saved by lossless object cleanup.</summary>
    public long EstimatedSavingsBytes { get; }

    /// <summary>Estimated lossless savings divided by the total parsed stream bytes.</summary>
    public double EstimatedSavingsRatio {
        get {
            long total = TotalStreamBytes;
            return total == 0 ? 0 : Math.Min(1, (double)Math.Max(0, EstimatedSavingsBytes) / total);
        }
    }

    /// <summary>Estimated lossless savings percentage across parsed stream bytes.</summary>
    public double EstimatedSavingsPercent => EstimatedSavingsRatio * 100d;

    /// <summary>Total stream count.</summary>
    public int StreamCount => Diagnostics.StreamCount;

    /// <summary>Total retained bytes across parsed streams.</summary>
    public long TotalStreamBytes => Diagnostics.Streams.Sum(static stream => stream.Length);

    /// <summary>Number of duplicate stream candidate groups.</summary>
    public int DuplicateStreamGroupCount => DuplicateStreams.Count;

    /// <summary>Number of stream objects participating in duplicate stream candidate groups.</summary>
    public int DuplicateStreamObjectCount => DuplicateStreams.Sum(static group => group.ObjectNumbers.Count);

    /// <summary>True when duplicate stream candidates were found.</summary>
    public bool HasDuplicateStreams => DuplicateStreamGroupCount > 0;

    /// <summary>Largest retained stream length in bytes.</summary>
    public long LargestStreamBytes => LargestStreams.Count == 0 ? 0 : LargestStreams[0].Length;

    /// <summary>Number of optimization findings.</summary>
    public int FindingCount => Findings.Count;

    /// <summary>True when the report found any concrete optimization opportunity.</summary>
    public bool HasOptimizationOpportunities => EstimatedSavingsBytes > 0 || FindingCount > 0 || HasDuplicateStreams;

    /// <summary>Image stream count.</summary>
    public int ImageStreamCount {
        get {
            int count = 0;
            foreach (PdfStreamDiagnostic stream in Diagnostics.Streams) {
                if (string.Equals(stream.Subtype, "Image", StringComparison.Ordinal)) {
                    count++;
                }
            }

            return count;
        }
    }
}

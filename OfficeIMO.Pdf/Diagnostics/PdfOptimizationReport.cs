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

    /// <summary>Total stream count.</summary>
    public int StreamCount => Diagnostics.StreamCount;

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

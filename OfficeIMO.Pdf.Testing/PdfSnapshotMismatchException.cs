namespace OfficeIMO.Pdf.Testing;

/// <summary>
/// Thrown when a PDF does not match its approved snapshot or no snapshot has been approved yet.
/// Every test framework reports it as a failure.
/// </summary>
public sealed class PdfSnapshotMismatchException : Exception {
    internal PdfSnapshotMismatchException(string snapshotName, string message, IReadOnlyList<string> receivedFiles, string? reviewPath)
        : base(message) {
        SnapshotName = snapshotName;
        ReceivedFiles = receivedFiles;
        ReviewPath = reviewPath;
    }

    /// <summary>Snapshot name passed to <see cref="PdfSnapshot.Verify(byte[], string, PdfSnapshotOptions?, string)"/>.</summary>
    public string SnapshotName { get; }

    /// <summary>Full paths of the <c>.received.</c> files written for review, for example to attach them to a test result.</summary>
    public IReadOnlyList<string> ReceivedFiles { get; }

    /// <summary>Full path of the HTML review gallery, or null when there was no approved snapshot to compare against.</summary>
    public string? ReviewPath { get; }
}

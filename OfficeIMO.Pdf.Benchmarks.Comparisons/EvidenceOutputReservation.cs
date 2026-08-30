namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed class EvidenceOutputReservation : IDisposable {
    private const string LockFileName = ".html-pdf-evidence.lock";
    private readonly string _lockPath;
    private FileStream? _lockStream;

    private EvidenceOutputReservation(string lockPath, FileStream lockStream) {
        _lockPath = lockPath;
        _lockStream = lockStream;
    }

    internal static EvidenceOutputReservation Acquire(string outputDirectory) {
        Directory.CreateDirectory(outputDirectory);
        string lockPath = Path.Combine(outputDirectory, LockFileName);
        FileStream lockStream;
        try {
            lockStream = new FileStream(lockPath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        } catch (IOException exception) {
            throw new IOException(
                $"HTML-to-PDF evidence output is already reserved or contains a stale reservation: '{outputDirectory}'.",
                exception);
        }

        var reservation = new EvidenceOutputReservation(lockPath, lockStream);
        if (Directory.EnumerateFileSystemEntries(outputDirectory)
            .Any(path => !string.Equals(Path.GetFileName(path), LockFileName, StringComparison.Ordinal))) {
            reservation.Dispose();
            throw new IOException(
                $"HTML-to-PDF evidence output must be a new or empty directory: '{outputDirectory}'.");
        }
        return reservation;
    }

    public void Dispose() {
        FileStream? stream = Interlocked.Exchange(ref _lockStream, null);
        if (stream == null) return;
        stream.Dispose();
        try {
            File.Delete(_lockPath);
        } catch (FileNotFoundException) {
            // The reservation is already absent.
        }
    }
}

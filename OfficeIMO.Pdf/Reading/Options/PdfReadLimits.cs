namespace OfficeIMO.Pdf;

/// <summary>Resource budgets applied while parsing PDF syntax and object graphs.</summary>
public sealed class PdfReadLimits {
    internal const int DefaultMaxDecodedStreamBytes = 256 * 1024 * 1024;

    /// <summary>Maximum input byte count accepted before text/object scanning. Default: 512 MiB.</summary>
    public long MaxInputBytes { get; set; } = 512L * 1024L * 1024L;

    /// <summary>Maximum number of indirect object declarations accepted. Default: 500,000.</summary>
    public int MaxIndirectObjects { get; set; } = 500_000;

    /// <summary>Maximum raw byte count allocated for one stream. Default: 256 MiB.</summary>
    public int MaxRawStreamBytes { get; set; } = 256 * 1024 * 1024;

    /// <summary>Maximum decoded byte count produced from one filtered stream. Default: 256 MiB.</summary>
    public int MaxDecodedStreamBytes { get; set; } = DefaultMaxDecodedStreamBytes;

    /// <summary>Maximum wall-clock time spent in the core object parsing pass. Default: 30 seconds.</summary>
    public TimeSpan MaxObjectParsingTime { get; set; } = TimeSpan.FromSeconds(30);

    internal void Validate() {
        if (MaxInputBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxInputBytes), MaxInputBytes, "Maximum input bytes must be positive.");
        }

        if (MaxIndirectObjects <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxIndirectObjects), MaxIndirectObjects, "Maximum indirect objects must be positive.");
        }

        if (MaxRawStreamBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxRawStreamBytes), MaxRawStreamBytes, "Maximum raw stream bytes must be positive.");
        }

        if (MaxDecodedStreamBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxDecodedStreamBytes), MaxDecodedStreamBytes, "Maximum decoded stream bytes must be positive.");
        }

        if (MaxObjectParsingTime <= TimeSpan.Zero) {
            throw new ArgumentOutOfRangeException(nameof(MaxObjectParsingTime), MaxObjectParsingTime, "Maximum object parsing time must be positive.");
        }
    }
}

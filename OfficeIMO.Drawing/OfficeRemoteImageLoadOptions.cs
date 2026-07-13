using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Controls bounded HTTP and HTTPS image retrieval.
/// </summary>
public sealed class OfficeRemoteImageLoadOptions {
    /// <summary>Default maximum response size: 10 MiB.</summary>
    public const long DefaultMaximumBytes = 10L * 1024L * 1024L;

    /// <summary>Maximum response size in bytes.</summary>
    public long MaximumBytes { get; set; } = DefaultMaximumBytes;

    /// <summary>Maximum time allowed for the complete request.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(30);

    /// <summary>Maximum number of same-origin redirects that may be followed.</summary>
    public int MaximumRedirects { get; set; } = 5;

    internal Snapshot CreateSnapshot() {
        if (MaximumBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaximumBytes), "MaximumBytes must be greater than zero.");
        }

        if (Timeout <= TimeSpan.Zero && Timeout != System.Threading.Timeout.InfiniteTimeSpan) {
            throw new ArgumentOutOfRangeException(nameof(Timeout), "Timeout must be positive or infinite.");
        }

        if (MaximumRedirects < 0) {
            throw new ArgumentOutOfRangeException(nameof(MaximumRedirects), "MaximumRedirects cannot be negative.");
        }

        return new Snapshot(MaximumBytes, Timeout, MaximumRedirects);
    }

    internal sealed class Snapshot {
        internal Snapshot(long maximumBytes, TimeSpan timeout, int maximumRedirects) {
            MaximumBytes = maximumBytes;
            Timeout = timeout;
            MaximumRedirects = maximumRedirects;
        }

        internal long MaximumBytes { get; }
        internal TimeSpan Timeout { get; }
        internal int MaximumRedirects { get; }
    }
}

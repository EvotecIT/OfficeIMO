using System;
using System.Collections.Generic;

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

    /// <summary>
    /// Allows loopback, private, link-local, multicast, and otherwise non-public destination
    /// addresses. The secure default is <c>false</c>; enable this only for explicitly trusted
    /// intranet image sources.
    /// </summary>
    public bool AllowPrivateNetworkAddresses { get; set; }

    /// <summary>
    /// Optional exact host allowlist. When non-empty, every initial and redirected destination
    /// must match one of these host names (case-insensitive).
    /// </summary>
    public ISet<string> AllowedHosts { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

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

        var allowedHosts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string host in AllowedHosts) {
            if (string.IsNullOrWhiteSpace(host)) {
                throw new ArgumentException("AllowedHosts cannot contain an empty host name.", nameof(AllowedHosts));
            }
            allowedHosts.Add(host.Trim().TrimEnd('.'));
        }

        return new Snapshot(MaximumBytes, Timeout, MaximumRedirects, AllowPrivateNetworkAddresses, allowedHosts);
    }

    internal sealed class Snapshot {
        internal Snapshot(long maximumBytes, TimeSpan timeout, int maximumRedirects, bool allowPrivateNetworkAddresses, HashSet<string> allowedHosts) {
            MaximumBytes = maximumBytes;
            Timeout = timeout;
            MaximumRedirects = maximumRedirects;
            AllowPrivateNetworkAddresses = allowPrivateNetworkAddresses;
            AllowedHosts = allowedHosts;
        }

        internal long MaximumBytes { get; }
        internal TimeSpan Timeout { get; }
        internal int MaximumRedirects { get; }
        internal bool AllowPrivateNetworkAddresses { get; }
        internal HashSet<string> AllowedHosts { get; }
    }
}

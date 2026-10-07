namespace OfficeIMO.Html;

/// <summary>Input limits and deterministic HTML entry selection for a ZIP site bundle.</summary>
public sealed class HtmlSiteBundleOptions {
    /// <summary>Optional exact, case-sensitive HTML entry path. When absent, root index.html,
    /// then root index.htm, then a single HTML entry is selected; ambiguous bundles are rejected.</summary>
    public string? EntryPath { get; set; }

    /// <summary>Virtual root URI for archive paths. It must be an absolute HTTP(S) directory URI
    /// without credentials, query or fragment. It does not authorize network access.</summary>
    public Uri ArchiveBaseUri { get; set; } = new Uri("https://bundle.officeimo.invalid/");

    /// <summary>Maximum encoded input ZIP bytes.</summary>
    public long MaximumArchiveBytes { get; set; } = 64L * 1024L * 1024L;
    /// <summary>Maximum ZIP entries, including directory entries, checked before ZIP metadata is materialized.</summary>
    public int MaximumEntryCount { get; set; } = 2048;
    /// <summary>Maximum decoded bytes of any single entry.</summary>
    public long MaximumEntryBytes { get; set; } = 16L * 1024L * 1024L;
    /// <summary>Maximum combined decoded bytes of all entries, including unused resources.</summary>
    public long MaximumTotalDecodedBytes { get; set; } = 128L * 1024L * 1024L;
    /// <summary>Maximum declared expansion ratio for any nonempty entry.</summary>
    public double MaximumCompressionRatio { get; set; } = 200D;

    internal HtmlSiteBundleOptions Snapshot() {
        if (MaximumArchiveBytes < 1 || MaximumArchiveBytes > int.MaxValue)
            throw new ArgumentOutOfRangeException(nameof(MaximumArchiveBytes));
        if (MaximumEntryCount < 1) throw new ArgumentOutOfRangeException(nameof(MaximumEntryCount));
        if (MaximumEntryBytes < 1 || MaximumEntryBytes > int.MaxValue)
            throw new ArgumentOutOfRangeException(nameof(MaximumEntryBytes));
        if (MaximumTotalDecodedBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaximumTotalDecodedBytes));
        if (MaximumCompressionRatio < 1D || double.IsInfinity(MaximumCompressionRatio) || double.IsNaN(MaximumCompressionRatio))
            throw new ArgumentOutOfRangeException(nameof(MaximumCompressionRatio));
        if (ArchiveBaseUri == null || !ArchiveBaseUri.IsAbsoluteUri
            || (ArchiveBaseUri.Scheme != Uri.UriSchemeHttps && ArchiveBaseUri.Scheme != Uri.UriSchemeHttp)
            || ArchiveBaseUri.UserInfo.Length != 0 || ArchiveBaseUri.Query.Length != 0 || ArchiveBaseUri.Fragment.Length != 0
            || !ArchiveBaseUri.AbsolutePath.EndsWith("/", StringComparison.Ordinal)) {
            throw new ArgumentException("The archive base URI must be an absolute HTTP(S) directory URI without credentials, query or fragment.", nameof(ArchiveBaseUri));
        }
        return new HtmlSiteBundleOptions {
            EntryPath = EntryPath,
            ArchiveBaseUri = ArchiveBaseUri,
            MaximumArchiveBytes = MaximumArchiveBytes,
            MaximumEntryCount = MaximumEntryCount,
            MaximumEntryBytes = MaximumEntryBytes,
            MaximumTotalDecodedBytes = MaximumTotalDecodedBytes,
            MaximumCompressionRatio = MaximumCompressionRatio
        };
    }
}

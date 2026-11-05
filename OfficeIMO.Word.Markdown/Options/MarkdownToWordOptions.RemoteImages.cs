using OfficeIMO.Drawing;

namespace OfficeIMO.Word.Markdown;

public partial class MarkdownWordExportOptions {
    /// <summary>Downloads allowed HTTP(S) images asynchronously before conversion when no caller resolver is supplied. Disabled by default.</summary>
    public bool AllowRemoteImages { get; set; }

    /// <summary>Network, redirect and byte limits for opt-in image retrieval. The shared loader blocks private networks by default.</summary>
    /// <remarks>The smaller of MaximumBytes and MaximumRemoteImageBytes applies. A caller RemoteImageResolver continues to own its own I/O policy.</remarks>
    public OfficeRemoteImageLoadOptions? RemoteImageLoadOptions { get; set; }

    internal static OfficeRemoteImageLoadOptions? CopyRemoteImageOptions(OfficeRemoteImageLoadOptions? source) {
        if (source == null) return null;
        var copy = new OfficeRemoteImageLoadOptions { MaximumBytes = source.MaximumBytes, Timeout = source.Timeout,
            MaximumRedirects = source.MaximumRedirects, AllowPrivateNetworkAddresses = source.AllowPrivateNetworkAddresses };
        foreach (string host in source.AllowedHosts) copy.AllowedHosts.Add(host);
        return copy;
    }
}

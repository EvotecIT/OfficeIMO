namespace OfficeIMO.Reader.Web;

/// <summary>Creates explicit web transports for existing Reader instances.</summary>
public static class OfficeDocumentReaderWebExtensions {
    /// <summary>
    /// Creates a bounded web reader that downloads through OfficeIMO's address-checked transport, without mutating
    /// the source Reader.
    /// </summary>
    /// <remarks>
    /// Loopback, private, link-local and other non-public addresses are refused after name resolution and on every
    /// redirect unless <see cref="ReaderWebOptions.AllowPrivateNetworkAddresses"/> is set.
    /// </remarks>
    public static OfficeDocumentWebReader CreateWebReader(
        this OfficeDocumentReader reader,
        ReaderWebOptions? options = null) {
        return new OfficeDocumentWebReader(reader, options);
    }

    /// <summary>
    /// Creates a bounded web reader without mutating the source Reader or taking ownership of the HTTP client.
    /// </summary>
    /// <remarks>
    /// The supplied client's handler resolves names and connects, so Reader Web screens only IP literals and
    /// localhost names for it. For untrusted URIs, prefer the overload without a client, or configure the handler to
    /// check resolved addresses at connection time and not to follow redirects itself.
    /// </remarks>
    public static OfficeDocumentWebReader CreateWebReader(
        this OfficeDocumentReader reader,
        HttpClient httpClient,
        ReaderWebOptions? options = null) {
        return new OfficeDocumentWebReader(reader, httpClient, options);
    }
}

namespace OfficeIMO.Reader.Web;

/// <summary>Creates explicit web transports for existing Reader instances.</summary>
public static class OfficeDocumentReaderWebExtensions {
    /// <summary>
    /// Creates a bounded web reader without mutating the source Reader or taking ownership of the HTTP client.
    /// </summary>
    public static OfficeDocumentWebReader CreateWebReader(
        this OfficeDocumentReader reader,
        HttpClient httpClient,
        ReaderWebOptions? options = null) {
        return new OfficeDocumentWebReader(reader, httpClient, options);
    }
}

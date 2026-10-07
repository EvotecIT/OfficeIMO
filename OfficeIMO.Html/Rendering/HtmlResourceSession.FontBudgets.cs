namespace OfficeIMO.Html;

public sealed partial class HtmlResourceSession {
    /// <summary>Accepted decoded font bytes retained by the operation, including nested
    /// documents and installed fallback faces. Encoded resource bytes remain separate.</summary>
    public long DecodedFontBytes { get; private set; }

    internal long RemainingFontBytes => MaxTotalResourceBytes - AcceptedResourceBytes - DecodedFontBytes;

    internal void AcceptDecodedFontBytes(int bytes) {
        if (bytes < 0 || bytes > RemainingFontBytes)
            throw new InvalidOperationException("Decoded font data exceeds the operation-wide resource budget.");
        DecodedFontBytes += bytes;
    }
}

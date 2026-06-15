namespace OfficeIMO.Rtf.Html;

/// <summary>
/// Converts between semantic HTML and the dependency-free OfficeIMO RTF document model.
/// </summary>
public static class RtfHtmlConverter {
    /// <summary>Converts semantic HTML to an RTF document model.</summary>
    public static RtfDocument FromHtml(string html, RtfHtmlReadOptions? options = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        return RtfHtmlReader.Read(html, options ?? new RtfHtmlReadOptions());
    }

    /// <summary>Converts an RTF document model to semantic HTML.</summary>
    public static string ToHtml(RtfDocument document, RtfHtmlSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        return RtfHtmlWriter.Write(document, options ?? new RtfHtmlSaveOptions());
    }
}

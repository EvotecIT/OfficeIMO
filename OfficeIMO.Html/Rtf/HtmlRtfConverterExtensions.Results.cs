namespace OfficeIMO.Html;

public static partial class HtmlRtfConverterExtensions {
    /// <summary>Imports HTML into RTF and returns structured conversion evidence.</summary>
    public static HtmlToRtfResult ToRtfDocumentResult(this string html, HtmlToRtfOptions? options = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        HtmlToRtfOptions resolved = (options ?? new HtmlToRtfOptions()).Clone();
        RtfDocument document = RtfHtmlReader.Read(html, resolved);
        return new HtmlToRtfResult(document, resolved.HtmlDiagnostics, resolved.Diagnostics.AsReadOnly(), resolved.ConversionReport);
    }

    /// <summary>Imports a prepared shared HTML document into RTF and returns structured evidence.</summary>
    public static HtmlToRtfResult ToRtfDocumentResult(this HtmlConversionDocument document, HtmlToRtfOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlToRtfOptions resolved = (options ?? new HtmlToRtfOptions()).Clone();
        RtfDocument rtfDocument = RtfHtmlReader.Read(document.DocumentForConversion, resolved);
        return new HtmlToRtfResult(rtfDocument, resolved.HtmlDiagnostics, resolved.Diagnostics.AsReadOnly(), resolved.ConversionReport);
    }

    /// <summary>Exports RTF to semantic HTML and returns structured conversion evidence.</summary>
    public static RtfToHtmlResult ToHtmlResult(this RtfDocument document, RtfToHtmlOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        RtfToHtmlOptions resolved = (options ?? new RtfToHtmlOptions()).Clone();
        string html = ToHtmlCore(document, resolved);
        return new RtfToHtmlResult(html, resolved.HtmlDiagnostics, resolved.Diagnostics.AsReadOnly(), resolved.ConversionReport);
    }
}

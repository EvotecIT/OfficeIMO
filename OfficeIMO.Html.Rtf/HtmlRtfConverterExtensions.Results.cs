namespace OfficeIMO.Html;

public static partial class HtmlRtfConverterExtensions {
    /// <summary>Imports a prepared shared HTML document into RTF and returns structured evidence.</summary>
    public static HtmlToRtfResult ToRtfDocumentResult(this HtmlConversionDocument document, HtmlToRtfOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlToRtfOptions resolved = (options ?? new HtmlToRtfOptions()).Clone();
        HtmlUrlPolicy requestedHyperlinkPolicy = resolved.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile();
        HtmlUrlPolicy requestedResourcePolicy = resolved.ResourceUrlPolicy ?? requestedHyperlinkPolicy;
        resolved.UrlPolicy = HtmlUrlPolicy.Intersect(document.HyperlinkUrlPolicy, requestedHyperlinkPolicy);
        resolved.ResourceUrlPolicy = HtmlUrlPolicy.Intersect(document.ResourceUrlPolicy, requestedResourcePolicy);
        if (resolved.BaseUri == null) resolved.BaseUri = document.BaseUri;
        AngleSharp.Html.Dom.IHtmlDocument sourceDocument = document.CreateSourceDocumentForConversion();
        HtmlNormalizer.SanitizePreparedDocumentStructure(sourceDocument);
        HtmlActiveMediaFilter.Filter(sourceDocument, document.MediaContext);
        RtfDocument rtfDocument = RtfHtmlReader.Read(sourceDocument, resolved);
        return new HtmlToRtfResult(
            rtfDocument,
            document.Diagnostics.Concat(document.ResourceManifest.Diagnostics).Concat(resolved.HtmlDiagnostics),
            resolved.Diagnostics.AsReadOnly(),
            resolved.ConversionReport);
    }

    /// <summary>Exports RTF to semantic HTML and returns structured conversion evidence.</summary>
    public static RtfToHtmlResult ToHtmlResult(this RtfDocument document, RtfToHtmlOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        RtfToHtmlOptions resolved = (options ?? new RtfToHtmlOptions()).Clone();
        string html = ToHtmlCore(document, resolved);
        return new RtfToHtmlResult(html, resolved.HtmlDiagnostics, resolved.Diagnostics.AsReadOnly(), resolved.ConversionReport);
    }
}

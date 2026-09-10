using AngleSharp;

namespace OfficeIMO.Html;

public sealed partial class HtmlConversionDocument {
    /// <summary>
    /// Serializes the source document with its effective absolute base URI so relative references
    /// still resolve against the original location when the HTML is saved elsewhere.
    /// </summary>
    /// <remarks>
    /// This preserves source markup rather than applying conversion normalization or sanitization.
    /// Referenced resources are not embedded or copied. Use <see cref="SourceHtml"/> for the exact
    /// original text, or <see cref="HtmlForConversion"/> for policy-normalized conversion HTML.
    /// </remarks>
    public string ExportSourceHtml() {
        if (BaseUri == null || !BaseUri.IsAbsoluteUri) return SourceHtml;
        var document = CreateSourceDocumentForConversion();
        var baseElement = document.QuerySelector("base[href]");
        if (baseElement == null) {
            baseElement = document.CreateElement("base");
            document.Head!.InsertBefore(baseElement, document.Head.FirstChild);
        }
        baseElement.SetAttribute("href", BaseUri.AbsoluteUri);
        return document.ToHtml();
    }
}

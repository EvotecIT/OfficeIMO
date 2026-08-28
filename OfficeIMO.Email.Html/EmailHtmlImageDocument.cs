using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Email;

/// <summary>One image source discovered in an email HTML body.</summary>
public sealed class EmailHtmlImageReference {
    internal EmailHtmlImageReference(int index, string source) {
        Index = index;
        Source = source;
    }

    /// <summary>Zero-based image index in document order.</summary>
    public int Index { get; }

    /// <summary>Decoded value of the image's <c>src</c> attribute.</summary>
    public string Source { get; }
}

/// <summary>
/// Bounded, network-free DOM view used to discover and rewrite image sources in email HTML.
/// </summary>
/// <remarks>
/// Parsing and HTML recovery are owned by OfficeIMO.Html. This type never resolves or downloads a
/// resource; callers remain responsible for authorizing local files and remote network access.
/// </remarks>
public sealed class EmailHtmlImageDocument {
    private readonly IHtmlDocument _document;
    private readonly IElement[] _images;
    private readonly bool _preserveDocumentEnvelope;

    private EmailHtmlImageDocument(IHtmlDocument document, bool preserveDocumentEnvelope) {
        _document = document;
        _images = document.QuerySelectorAll("img[src]").ToArray();
        _preserveDocumentEnvelope = preserveDocumentEnvelope;
        Images = _images
            .Select((image, index) => new EmailHtmlImageReference(index, image.GetAttribute("src") ?? string.Empty))
            .ToArray();
    }

    /// <summary>Images with explicit <c>src</c> attributes in document order.</summary>
    public IReadOnlyList<EmailHtmlImageReference> Images { get; }

    /// <summary>
    /// Parses an HTML fragment or document under the shared OfficeIMO limits without performing network access.
    /// </summary>
    /// <param name="html">HTML fragment or document to inspect.</param>
    public static EmailHtmlImageDocument Parse(string html) {
        if (html == null) throw new ArgumentNullException(nameof(html));

        HtmlUrlPolicy compatibleResources = HtmlUrlPolicy.CreateOfficeIMOProfile();
        var options = HtmlConversionDocumentOptions.CreateUntrustedProfile();
        options.IncludeNormalizedHtml = false;
        options.UseBodyContentsOnly = false;
        options.ResourceUrlPolicy = compatibleResources;
        HtmlConversionDocument conversion = HtmlConversionDocument.Parse(html, options);
        IHtmlDocument document = conversion.CreateDocumentForConversion();
        bool preserveEnvelope = ContainsDocumentEnvelope(html);
        return new EmailHtmlImageDocument(document, preserveEnvelope);
    }

    /// <summary>Rewrites one image source selected by its stable document-order index.</summary>
    /// <param name="imageIndex">Index from <see cref="EmailHtmlImageReference.Index"/>.</param>
    /// <param name="source">Replacement source, such as a <c>cid:</c> reference.</param>
    public void SetImageSource(int imageIndex, string source) {
        if (imageIndex < 0 || imageIndex >= _images.Length) throw new ArgumentOutOfRangeException(nameof(imageIndex));
        if (string.IsNullOrWhiteSpace(source)) throw new ArgumentException("An image source is required.", nameof(source));
        _images[imageIndex].SetAttribute("src", source);
    }

    /// <summary>Serializes the current DOM while retaining fragment versus document shape.</summary>
    public string ToHtml() {
        if (_preserveDocumentEnvelope) return _document.DocumentElement?.OuterHtml ?? string.Empty;
        return _document.Body?.InnerHtml ?? _document.DocumentElement?.InnerHtml ?? string.Empty;
    }

    private static bool ContainsDocumentEnvelope(string html) =>
        html.IndexOf("<!doctype", StringComparison.OrdinalIgnoreCase) >= 0 ||
        html.IndexOf("<html", StringComparison.OrdinalIgnoreCase) >= 0 ||
        html.IndexOf("<head", StringComparison.OrdinalIgnoreCase) >= 0 ||
        html.IndexOf("<body", StringComparison.OrdinalIgnoreCase) >= 0;
}

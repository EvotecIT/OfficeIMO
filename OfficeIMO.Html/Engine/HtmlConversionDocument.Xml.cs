using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Html;

public sealed partial class HtmlConversionDocument {
    // Explicit package boundary: never infer XML semantics from ordinary HTML source.
    internal static HtmlConversionDocument ParseXhtml(string source, HtmlConversionDocumentOptions options,
        CancellationToken token = default) {
        HtmlConversionDocumentOptions resolved = options.Clone();
        resolved.Validate();
        token.ThrowIfCancellationRequested();
        HtmlConversionInputGuard.ValidateSource(source, resolved.Limits);
        using var text = new StringReader(source);
        using var reader = XmlReader.Create(text, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Ignore, XmlResolver = null,
            MaxCharactersInDocument = resolved.Limits.MaxInputCharacters ?? 0
        });
        XDocument xml = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        if (xml.Root?.Name != XName.Get("html", Dom.HtmlElement.HtmlNamespace))
            throw new InvalidDataException("XHTML conversion requires an html root in the XHTML namespace.");
        var document = HtmlXmlDocumentParser.CreateDocument(xml, resolved.Limits, token);
        return new HtmlConversionDocument(source, document, resolved,
            HtmlDocumentParser.ResolveEffectiveBaseUri(document, resolved.BaseUri));
    }
}

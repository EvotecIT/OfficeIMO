using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Pdf;

internal static class PdfXmpMetadataSynchronizer {
    private static readonly char[] KeywordSeparators = { ',', ';' };
    private const string DublinCoreNamespaceUri = "http://purl.org/dc/elements/1.1/";
    private const string PdfNamespaceUri = "http://ns.adobe.com/pdf/1.3/";
    private const string RdfNamespaceUri = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";

    internal static byte[] Synchronize(string rawXml, PdfMetadata metadata) {
        Guard.NotNull(rawXml, nameof(rawXml));
        Guard.NotNull(metadata, nameof(metadata));

        XDocument document = Parse(rawXml);
        XNamespace rdf = RdfNamespaceUri;
        XElement description = document
            .Descendants(rdf + "Description")
            .FirstOrDefault() ?? AddDescription(document, rdf);

        ReplaceAlt(description, "title", metadata.Title);
        ReplaceSequence(description, "creator", metadata.Author);
        ReplaceAlt(description, "description", metadata.Subject);
        ReplaceSubjectBag(description, metadata.Keywords);
        ReplaceText(description, PdfNamespaceUri, "Keywords", metadata.Keywords);
        if (!description.Elements().Any(static element =>
            element.Name.LocalName == "Producer" &&
            element.Name.NamespaceName == PdfNamespaceUri)) {
            description.Add(new XElement(XName.Get("Producer", PdfNamespaceUri), "OfficeIMO.Pdf"));
        }

        return Serialize(document);
    }

    private static XDocument Parse(string rawXml) {
        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                MaxCharactersInDocument = PdfReadDocument.MaxXmpMetadataBytes,
                XmlResolver = null
            };
            using var textReader = new StringReader(rawXml);
            using XmlReader reader = XmlReader.Create(textReader, settings);
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        } catch (XmlException ex) {
            throw new InvalidOperationException("The existing XMP packet is not well-formed XML and cannot be synchronized without discarding custom metadata.", ex);
        }
    }

    private static XElement AddDescription(XDocument document, XNamespace rdf) {
        XElement? rdfRoot = document.Descendants(rdf + "RDF").FirstOrDefault();
        if (rdfRoot is null) {
            throw new InvalidOperationException("The existing XMP packet does not contain an RDF root and cannot be synchronized without discarding custom metadata.");
        }

        var description = new XElement(rdf + "Description", new XAttribute(rdf + "about", string.Empty));
        rdfRoot.Add(description);
        return description;
    }

    private static void ReplaceAlt(XElement description, string localName, string? value) {
        Remove(description, DublinCoreNamespaceUri, localName);
        if (string.IsNullOrEmpty(value)) {
            return;
        }

        XNamespace dc = DublinCoreNamespaceUri;
        XNamespace rdf = RdfNamespaceUri;
        description.Add(new XElement(
            dc + localName,
            new XElement(
                rdf + "Alt",
                new XElement(rdf + "li", new XAttribute(XNamespace.Xml + "lang", "x-default"), value))));
    }

    private static void ReplaceSequence(XElement description, string localName, string? value) {
        Remove(description, DublinCoreNamespaceUri, localName);
        if (string.IsNullOrEmpty(value)) {
            return;
        }

        XNamespace dc = DublinCoreNamespaceUri;
        XNamespace rdf = RdfNamespaceUri;
        description.Add(new XElement(dc + localName, new XElement(rdf + "Seq", new XElement(rdf + "li", value))));
    }

    private static void ReplaceSubjectBag(XElement description, string? keywords) {
        Remove(description, DublinCoreNamespaceUri, "subject");
        string[] values = SplitKeywords(keywords);
        if (values.Length == 0) {
            return;
        }

        XNamespace dc = DublinCoreNamespaceUri;
        XNamespace rdf = RdfNamespaceUri;
        description.Add(new XElement(
            dc + "subject",
            new XElement(rdf + "Bag", values.Select(value => new XElement(rdf + "li", value)))));
    }

    private static void ReplaceText(XElement description, string namespaceUri, string localName, string? value) {
        Remove(description, namespaceUri, localName);
        if (!string.IsNullOrEmpty(value)) {
            description.Add(new XElement(XName.Get(localName, namespaceUri), value));
        }
    }

    private static void Remove(XElement description, string namespaceUri, string localName) {
        description.Elements()
            .Where(element => element.Name.LocalName == localName && element.Name.NamespaceName == namespaceUri)
            .Remove();
    }

    private static string[] SplitKeywords(string? keywords) {
        if (string.IsNullOrWhiteSpace(keywords)) {
            return Array.Empty<string>();
        }

        return keywords!
            .Split(KeywordSeparators, StringSplitOptions.RemoveEmptyEntries)
            .Select(static value => value.Trim())
            .Where(static value => value.Length > 0)
            .ToArray();
    }

    private static byte[] Serialize(XDocument document) {
        using var output = new MemoryStream();
        var settings = new XmlWriterSettings {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            OmitXmlDeclaration = document.Declaration is null
        };
        using (XmlWriter writer = XmlWriter.Create(output, settings)) {
            document.WriteTo(writer);
        }

        return output.ToArray();
    }
}

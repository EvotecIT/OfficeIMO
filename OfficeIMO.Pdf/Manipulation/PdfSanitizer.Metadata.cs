using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

/// <summary>
/// User metadata removal keeps what identifies the document rather than the people and tools behind it: the title, which readers
/// and screen readers announce and PDF/UA requires, and the PDF/A and PDF/UA identification, without which a conforming file stops
/// claiming its conformance. The XMP packet is reduced to those properties plus the XMP equivalents of the technical Info entries
/// that stay, because PDF/A requires the two to agree.
/// </summary>
internal static partial class PdfSanitizer {
    private const string TitleInfoKey = "Title";
    private const string RdfNamespace = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
    private const string DublinCoreNamespace = "http://purl.org/dc/elements/1.1/";
    private const string PdfPropertiesNamespace = "http://ns.adobe.com/pdf/1.3/";
    private const string XmpBasicNamespace = "http://ns.adobe.com/xap/1.0/";
    private const string PdfAIdentificationNamespace = "http://www.aiim.org/pdfa/ns/id/";
    private const string PdfAExtensionNamespace = "http://www.aiim.org/pdfa/ns/extension/";
    private const string PdfASchemaNamespace = "http://www.aiim.org/pdfa/ns/schema#";

    internal const string RetainedTitleNote =
        "Kept because the title names the document for readers and screen readers, and accessible PDF (PDF/UA) requires it.";
    internal const string RetainedConformanceNote =
        "Kept because the file would otherwise stop declaring the PDF/A or PDF/UA conformance it was made for.";

    private static readonly HashSet<string> RetainedXmpProperties = new(StringComparer.Ordinal) {
        DublinCoreNamespace + "title",
        PdfPropertiesNamespace + "Producer",
        PdfPropertiesNamespace + "Trapped",
        XmpBasicNamespace + "CreateDate",
        XmpBasicNamespace + "ModifyDate"
    };

    private static readonly HashSet<string> RetainedXmpNamespaces = new(StringComparer.Ordinal) {
        PdfAIdentificationNamespace,
        PdfUaIdentification.NamespaceUri
    };

    /// <summary>What user metadata removal does with the catalog XMP packet.</summary>
    private sealed class PdfXmpReduction {
        /// <summary>The reduced packet, or null when nothing worth keeping is left and the stream is removed.</summary>
        internal byte[]? Packet { get; set; }

        /// <summary>Whether the packet holds anything that removal drops.</summary>
        internal bool RemovesProperties { get; set; }

        internal List<PdfSanitizationItem> Retained { get; } = new();
    }

    private static PdfXmpReduction ReduceXmpMetadata(Dictionary<int, PdfIndirectObject> objects, PdfObject metadata) {
        var reduction = new PdfXmpReduction { RemovesProperties = true };
        if (Resolve(objects, metadata) is not PdfStream stream ||
            !StreamDecoder.TryDecode(stream, PdfReadDocument.MaxXmpMetadataBytes, out byte[] data, objects)) {
            return reduction;
        }

        XDocument document;
        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = PdfReadDocument.MaxXmpMetadataBytes
            };
            using var reader = XmlReader.Create(new MemoryStream(data, writable: false), settings);
            document = XDocument.Load(reader);
        } catch (XmlException) {
            return reduction;
        }

        XNamespace rdf = RdfNamespace;
        XElement? root = document.Descendants(rdf + "RDF").FirstOrDefault();
        if (root is null) return reduction;

        bool removes = false;
        var attributes = new List<XAttribute>();
        var elements = new List<XElement>();
        var prefixes = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (XElement node in root.Elements()) {
            if (node.Name != rdf + "Description") {
                removes = true;
                continue;
            }
            foreach (XAttribute attribute in node.Attributes()) {
                if (attribute.IsNamespaceDeclaration || attribute.Name.Namespace == rdf || attribute.Name.NamespaceName.Length == 0) continue;
                if (IsRetainedXmpProperty(attribute.Name)) {
                    attributes.Add(new XAttribute(attribute));
                    RememberPrefix(node, attribute.Name.Namespace, prefixes);
                } else {
                    removes = true;
                }
            }
            foreach (XElement property in node.Elements()) {
                XElement? kept = property.Name == XName.Get("schemas", PdfAExtensionNamespace)
                    ? ReduceExtensionSchemas(property, ref removes)
                    : IsRetainedXmpProperty(property.Name) ? new XElement(property) : null;
                if (kept is null) {
                    removes = true;
                    continue;
                }
                elements.Add(kept);
                foreach (XElement descendant in property.DescendantsAndSelf()) {
                    RememberPrefix(descendant, descendant.Name.Namespace, prefixes);
                    foreach (XAttribute attribute in descendant.Attributes()) {
                        if (!attribute.IsNamespaceDeclaration) RememberPrefix(descendant, attribute.Name.Namespace, prefixes);
                    }
                }
            }
        }

        reduction.RemovesProperties = removes;
        AddRetainedXmpItems(attributes, elements, reduction.Retained);
        if (reduction.Retained.Count > 0) reduction.Packet = BuildRetainedXmp(attributes, elements, prefixes);
        return reduction;
    }

    private static bool IsRetainedXmpProperty(XName name) =>
        RetainedXmpNamespaces.Contains(name.NamespaceName) || RetainedXmpProperties.Contains(name.NamespaceName + name.LocalName);

    /// <summary>Keeps only the extension schema descriptions for namespaces the reduced packet still uses, such as pdfuaid in PDF/A-2 and PDF/A-3.</summary>
    private static XElement? ReduceExtensionSchemas(XElement schemas, ref bool removes) {
        var copy = new XElement(schemas);
        XNamespace rdf = RdfNamespace;
        List<XElement> entries = copy.Elements().Elements(rdf + "li").ToList();
        foreach (XElement entry in entries) {
            string? schemaNamespace = entry.Descendants(XName.Get("namespaceURI", PdfASchemaNamespace)).FirstOrDefault()?.Value.Trim();
            if (schemaNamespace is null || !RetainedXmpNamespaces.Contains(schemaNamespace)) {
                entry.Remove();
                removes = true;
            }
        }
        return copy.Elements().Elements(rdf + "li").Any() ? copy : null;
    }

    private static void RememberPrefix(XElement scope, XNamespace ns, Dictionary<string, string> prefixes) {
        if (ns == XNamespace.None || ns == XNamespace.Xml || prefixes.ContainsKey(ns.NamespaceName)) return;
        string? prefix = scope.GetPrefixOfNamespace(ns);
        if (!string.IsNullOrEmpty(prefix)) prefixes[ns.NamespaceName] = prefix!;
    }

    private static void AddRetainedXmpItems(List<XAttribute> attributes, List<XElement> elements, List<PdfSanitizationItem> retained) {
        XElement? title = elements.FirstOrDefault(static element => element.Name == XName.Get("title", DublinCoreNamespace));
        if (title is not null) {
            retained.Add(new PdfSanitizationItem(PdfSanitizationContentKind.UserMetadata, "Catalog/Metadata/dc:title", "XMP title",
                title.Descendants(XName.Get("li", RdfNamespace)).FirstOrDefault()?.Value ?? title.Value, retainedNote: RetainedTitleNote));
        }
        AddIdentification(PdfAIdentificationNamespace, "PDF/A", "XMP PDF/A identification", "Catalog/Metadata/pdfaid");
        AddIdentification(PdfUaIdentification.NamespaceUri, "PDF/UA", "XMP PDF/UA identification", "Catalog/Metadata/pdfuaid");

        void AddIdentification(string ns, string standard, string description, string location) {
            string? part = ReadIdentification(ns, "part");
            if (part is null) return;
            string value = standard + "-" + part + (ReadIdentification(ns, "conformance") ?? string.Empty);
            retained.Add(new PdfSanitizationItem(PdfSanitizationContentKind.UserMetadata, location, description, value, retainedNote: RetainedConformanceNote));
        }

        string? ReadIdentification(string ns, string localName) {
            XName name = XName.Get(localName, ns);
            string? value = elements.FirstOrDefault(element => element.Name == name)?.Value ??
                attributes.FirstOrDefault(attribute => attribute.Name == name)?.Value;
            return string.IsNullOrWhiteSpace(value) ? null : value!.Trim();
        }
    }

    private static byte[] BuildRetainedXmp(List<XAttribute> attributes, List<XElement> elements, Dictionary<string, string> prefixes) {
        XNamespace x = "adobe:ns:meta/";
        XNamespace rdf = RdfNamespace;
        var description = new XElement(rdf + "Description", new XAttribute(rdf + "about", string.Empty));
        foreach (KeyValuePair<string, string> prefix in prefixes.OrderBy(static item => item.Value, StringComparer.Ordinal)) {
            if (prefix.Key == RdfNamespace) continue;
            description.Add(new XAttribute(XNamespace.Xmlns + prefix.Value, prefix.Key));
        }
        foreach (XAttribute attribute in attributes) description.Add(attribute);
        foreach (XElement element in elements) description.Add(element);
        var packet = new XElement(x + "xmpmeta",
            new XAttribute(XNamespace.Xmlns + "x", x.NamespaceName),
            new XElement(rdf + "RDF", new XAttribute(XNamespace.Xmlns + "rdf", RdfNamespace), description));

        var settings = new XmlWriterSettings {
            OmitXmlDeclaration = true,
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = true,
            NewLineChars = "\n"
        };
        using var output = new MemoryStream();
        using (var writer = XmlWriter.Create(output, settings)) {
            writer.WriteProcessingInstruction("xpacket", "begin=\"﻿\" id=\"W5M0MpCehiHzreSzNTczkc9d\"");
            packet.WriteTo(writer);
            writer.WriteProcessingInstruction("xpacket", "end=\"w\"");
        }
        return output.ToArray();
    }

    /// <summary>Replaces the catalog XMP stream with <paramref name="packet"/>, uncompressed as PDF/A requires.</summary>
    private static bool TryReplaceMetadataStream(Dictionary<int, PdfIndirectObject> objects, PdfObject metadata, byte[] packet) {
        if (metadata is not PdfReference reference ||
            !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
            indirect.Value is not PdfStream) {
            return false;
        }
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("Metadata");
        dictionary.Items["Subtype"] = new PdfName("XML");
        objects[reference.ObjectNumber] = new PdfIndirectObject(indirect.ObjectNumber, indirect.Generation, new PdfStream(dictionary, packet));
        return true;
    }
}

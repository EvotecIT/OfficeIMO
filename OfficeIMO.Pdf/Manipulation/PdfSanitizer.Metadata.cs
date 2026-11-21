using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

/// <summary>
/// User metadata removal keeps what identifies the document rather than the people and tools behind it: the title, which readers
/// and screen readers announce and PDF/UA requires, and the PDF/A, PDF/UA and PDF/X identification, without which a conforming file
/// stops claiming its conformance. An e-invoice keeps its Factur-X or ZUGFeRD identification while its invoice data stays attached.
/// The XMP packet is reduced to those properties plus the XMP equivalents of the technical Info entries that stay, because PDF/A
/// requires the two to agree, and, for PDF/X, the document identifiers and metadata date PDF/X requires.
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
    private const string XmpMediaManagementNamespace = "http://ns.adobe.com/xap/1.0/mm/";
    private const string PdfXVersionInfoKey = "GTS_PDFXVersion";

    internal const string RetainedTitleNote =
        "Kept because the title names the document for readers and screen readers, and accessible PDF (PDF/UA) requires it.";
    internal const string RetainedConformanceNote =
        "Kept because the file would otherwise stop declaring the PDF/A, PDF/UA or PDF/X conformance it was made for.";
    internal const string RetainedPdfXIdentifiersNote =
        "Kept because PDF/X requires the document and instance IDs and the metadata date alongside its identification.";
    internal const string RetainedInvoiceNote =
        "Kept because it identifies the e-invoice data attached to this file; without it the invoice no longer reads as Factur-X or ZUGFeRD. Remove attachments too to drop both.";

    /// <summary>Info entries that declare conformance rather than describe people or tools, reported as retained.</summary>
    private static readonly HashSet<string> ConformanceInfoKeys = new(StringComparer.Ordinal) {
        PdfXVersionInfoKey, "GTS_PDFXConformance"
    };

    private static readonly string[] AlwaysRetainedXmpProperties = {
        DublinCoreNamespace + "title",
        PdfPropertiesNamespace + "Producer",
        PdfPropertiesNamespace + "Trapped",
        XmpBasicNamespace + "CreateDate",
        XmpBasicNamespace + "ModifyDate"
    };

    private static readonly string[] PdfXRetainedXmpProperties = {
        XmpBasicNamespace + "MetadataDate",
        XmpMediaManagementNamespace + "DocumentID",
        XmpMediaManagementNamespace + "InstanceID",
        XmpMediaManagementNamespace + "VersionID",
        XmpMediaManagementNamespace + "RenditionClass"
    };

    /// <summary>Factur-X, ZUGFeRD 2 and ZUGFeRD 1 identification namespaces.</summary>
    private static readonly string[] ElectronicInvoiceNamespaces = {
        PdfElectronicInvoiceMetadata.FacturXNamespaceUri,
        "urn:zugferd:pdfa:CrossIndustryDocument:invoice:2p0#",
        "urn:ferd:pdfa:CrossIndustryDocument:invoice:1p0#"
    };

    /// <summary>The XMP properties and namespaces one reduction keeps, which depend on what the file declares and what else is removed.</summary>
    private sealed class XmpRetention {
        internal HashSet<string> Properties { get; } = new(StringComparer.Ordinal);
        internal HashSet<string> Namespaces { get; } = new(StringComparer.Ordinal) { PdfAIdentificationNamespace, PdfUaIdentification.NamespaceUri };
        internal bool KeepsPdfX { get; set; }

        internal bool Keeps(XName name) => Namespaces.Contains(name.NamespaceName) || Properties.Contains(name.NamespaceName + name.LocalName);
    }

    /// <summary>What user metadata removal does with the catalog XMP packet.</summary>
    private sealed class PdfXmpReduction {
        /// <summary>The reduced packet, or null when nothing worth keeping is left and the stream is removed.</summary>
        internal byte[]? Packet { get; set; }

        /// <summary>Whether the packet holds anything that removal drops.</summary>
        internal bool RemovesProperties { get; set; }

        internal List<PdfSanitizationItem> Retained { get; } = new();
    }

    /// <summary>Works out what user metadata removal keeps from the catalog XMP packet and what it reports as retained.</summary>
    /// <param name="objects">The document's objects.</param>
    /// <param name="metadata">The catalog <c>/Metadata</c> value.</param>
    /// <param name="infoDeclaresPdfX">Whether the Info dictionary declares PDF/X, which keeps the PDF/X production properties even when XMP lacks pdfxid.</param>
    /// <param name="keepElectronicInvoice">Whether the embedded invoice data stays, so its Factur-X or ZUGFeRD identification must stay with it.</param>
    private static PdfXmpReduction ReduceXmpMetadata(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject metadata,
        bool infoDeclaresPdfX,
        bool keepElectronicInvoice) {
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

        var retention = new XmpRetention();
        foreach (string property in AlwaysRetainedXmpProperties) retention.Properties.Add(property);
        retention.KeepsPdfX = infoDeclaresPdfX || root.Descendants().Any(static element => element.Name.NamespaceName == PdfXIdentification.NamespaceUri) ||
            root.Descendants().Attributes().Any(static attribute => attribute.Name.NamespaceName == PdfXIdentification.NamespaceUri);
        if (retention.KeepsPdfX) {
            retention.Namespaces.Add(PdfXIdentification.NamespaceUri);
            foreach (string property in PdfXRetainedXmpProperties) retention.Properties.Add(property);
        }
        if (keepElectronicInvoice) {
            foreach (string invoiceNamespace in ElectronicInvoiceNamespaces) retention.Namespaces.Add(invoiceNamespace);
        }

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
                if (retention.Keeps(attribute.Name)) {
                    attributes.Add(new XAttribute(attribute));
                    RememberPrefix(node, attribute.Name.Namespace, prefixes);
                } else {
                    removes = true;
                }
            }
            foreach (XElement property in node.Elements()) {
                XElement? kept = property.Name == XName.Get("schemas", PdfAExtensionNamespace)
                    ? ReduceExtensionSchemas(property, retention, ref removes)
                    : retention.Keeps(property.Name) ? new XElement(property) : null;
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
        AddRetainedXmpItems(attributes, elements, retention, reduction.Retained);
        if (reduction.Retained.Count > 0) reduction.Packet = BuildRetainedXmp(attributes, elements, prefixes);
        return reduction;
    }

    /// <summary>Keeps only the extension schema descriptions for namespaces the reduced packet still uses, such as pdfuaid and fx in PDF/A-2 and PDF/A-3.</summary>
    private static XElement? ReduceExtensionSchemas(XElement schemas, XmpRetention retention, ref bool removes) {
        var copy = new XElement(schemas);
        XNamespace rdf = RdfNamespace;
        List<XElement> entries = copy.Elements().Elements(rdf + "li").ToList();
        foreach (XElement entry in entries) {
            string? schemaNamespace = entry.Descendants(XName.Get("namespaceURI", PdfASchemaNamespace)).FirstOrDefault()?.Value.Trim();
            if (schemaNamespace is null || !retention.Namespaces.Contains(schemaNamespace)) {
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

    private static void AddRetainedXmpItems(List<XAttribute> attributes, List<XElement> elements, XmpRetention retention, List<PdfSanitizationItem> retained) {
        XElement? title = elements.FirstOrDefault(static element => element.Name == XName.Get("title", DublinCoreNamespace));
        if (title is not null) {
            retained.Add(new PdfSanitizationItem(PdfSanitizationContentKind.UserMetadata, "Catalog/Metadata/dc:title", "XMP title",
                title.Descendants(XName.Get("li", RdfNamespace)).FirstOrDefault()?.Value ?? title.Value, retainedNote: RetainedTitleNote));
        }
        AddIdentification(PdfAIdentificationNamespace, "PDF/A", "XMP PDF/A identification", "Catalog/Metadata/pdfaid");
        AddIdentification(PdfUaIdentification.NamespaceUri, "PDF/UA", "XMP PDF/UA identification", "Catalog/Metadata/pdfuaid");
        string? pdfXVersion = ReadIdentification(PdfXIdentification.NamespaceUri, "GTS_PDFXVersion");
        if (pdfXVersion is not null) {
            retained.Add(new PdfSanitizationItem(PdfSanitizationContentKind.UserMetadata, "Catalog/Metadata/pdfxid", "XMP PDF/X identification",
                pdfXVersion, retainedNote: RetainedConformanceNote));
        }
        if (retention.KeepsPdfX) {
            string? documentId = ReadIdentification(XmpMediaManagementNamespace, "DocumentID");
            if (documentId is not null || ReadIdentification(XmpMediaManagementNamespace, "InstanceID") is not null) {
                retained.Add(new PdfSanitizationItem(PdfSanitizationContentKind.UserMetadata, "Catalog/Metadata/xmpMM", "XMP document and instance IDs",
                    documentId, retainedNote: RetainedPdfXIdentifiersNote));
            }
        }
        foreach (string invoiceNamespace in ElectronicInvoiceNamespaces) {
            if (!retention.Namespaces.Contains(invoiceNamespace)) continue;
            string? level = ReadIdentification(invoiceNamespace, "ConformanceLevel");
            string? fileName = ReadIdentification(invoiceNamespace, "DocumentFileName");
            if (level is null && fileName is null) continue;
            retained.Add(new PdfSanitizationItem(PdfSanitizationContentKind.UserMetadata, "Catalog/Metadata/invoice", "XMP e-invoice identification",
                level is null ? fileName : fileName is null ? level : fileName + " (" + level + ")", retainedNote: RetainedInvoiceNote));
        }

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

using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceXmp {
    private const string IptcNamespace = "http://iptc.org/std/Iptc4xmpExt/2008-02-29/";
    private static readonly XNamespace RdfNamespace = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
    private const string VocabularyPrefix = "http://cv.iptc.org/newscodes/digitalsourcetype/";

    internal static void Inspect(byte[] packet, OfficeProvenanceOptions options, OfficeProvenanceContext context, string location,
        bool carrierIsStructurallyValid = true,
        bool allowDirectRootIptc = false) {
        if (!TryLoad(packet, options, out XDocument? document) || document == null) return;
        int index = 0;
        foreach (XmpValue value in FindValues(document, allowDirectRootIptc)) {
            context.Add(new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.IptcDigitalSourceType,
                $"{location}/DigitalSourceType[{index++}]",
                isStructurallyValid: carrierIsStructurallyValid && value.IsStructurallyValid,
                value: value.Value,
                digitalSourceKind: value.Kind));
        }
    }

    internal static bool TryRemoveAiDeclarations(
        byte[] packet,
        OfficeProvenanceRemovalOptions options,
        string location,
        List<OfficeProvenanceChange> changes,
        out byte[] output,
        bool carrierIsStructurallyValid = true,
        bool allowDirectRootIptc = false) {
        output = packet;
        if (!carrierIsStructurallyValid && options.RequireStructurallyValidCarrier) return false;
        if (!options.RemoveAiSourceMetadata || !TryLoad(packet, options.Limits, out XDocument? document) || document == null) return false;
        XmpValue[] values = FindValues(document, allowDirectRootIptc).ToArray();
        bool changed = false;
        int index = 0;
        foreach (XmpValue value in values) {
            bool remove = value.Kind == OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia ||
                value.Kind == OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia;
            if (remove && (value.IsStructurallyValid || !options.RequireStructurallyValidCarrier)) {
                if (value.Element != null) value.Element.Remove();
                else value.Attribute?.Remove();
                changes.Add(new OfficeProvenanceChange(
                    OfficeProvenanceCarrierKind.IptcDigitalSourceType,
                    $"{location}/DigitalSourceType[{index}]",
                    removedBytes: 0));
                changed = true;
            }
            index++;
        }
        if (!changed) return false;
        using var stream = new OfficeProvenanceBoundedMemoryStream(
            options.EffectiveMaxOutputBytes,
            packet.Length);
        var settings = new XmlWriterSettings {
            Encoding = new UTF8Encoding(false),
            Indent = false,
            OmitXmlDeclaration = document.Declaration == null,
            NewLineHandling = NewLineHandling.None
        };
        using (XmlWriter writer = XmlWriter.Create(stream, settings)) document.Save(writer);
        output = stream.ToArray();
        return true;
    }

    /// <summary>Builds a minimal XMP packet that keeps only the digital source type declarations in <paramref name="packet"/>, or null when it declares none.</summary>
    internal static byte[]? CreateDigitalSourceTypePacket(byte[] packet, OfficeProvenanceOptions options) {
        if (!TryLoad(packet, options, out XDocument? document) || document == null) return null;
        string[] values = FindValues(document).Where(value => value.IsStructurallyValid && value.Value.Length > 0)
            .Select(value => value.Value).Distinct(StringComparer.Ordinal).ToArray();
        if (values.Length == 0) return null;
        XNamespace meta = "adobe:ns:meta/";
        XNamespace iptc = IptcNamespace;
        var root = new XElement(meta + "xmpmeta",
            new XAttribute(XNamespace.Xmlns + "x", meta.NamespaceName),
            new XElement(RdfNamespace + "RDF",
                new XAttribute(XNamespace.Xmlns + "rdf", RdfNamespace.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "Iptc4xmpExt", IptcNamespace),
                values.Select(value => new XElement(RdfNamespace + "Description",
                    new XAttribute(RdfNamespace + "about", string.Empty),
                    new XElement(iptc + "DigitalSourceType", new XAttribute(RdfNamespace + "resource", value))))));
        string text = "<?xpacket begin=\"\uFEFF\" id=\"W5M0MpCehiHzreSzNTczkc9d\"?>" +
            root.ToString(SaveOptions.DisableFormatting) + "<?xpacket end=\"w\"?>";
        return new UTF8Encoding(false).GetBytes(text);
    }

    /// <summary>Whether a packet carries digital source type declarations and nothing else.</summary>
    internal static bool HasOnlyDigitalSourceTypes(byte[] packet, OfficeProvenanceOptions options) {
        if (!TryLoad(packet, options, out XDocument? document) || document == null) return false;
        XElement[] graphs = document.Descendants(RdfNamespace + "RDF").ToArray();
        if (graphs.Length == 0) return false;
        bool found = false;
        foreach (XElement graph in graphs) {
            foreach (XElement description in graph.Elements()) {
                if (description.Name != RdfNamespace + "Description") return false;
                foreach (XAttribute attribute in description.Attributes()) {
                    if (attribute.IsNamespaceDeclaration || attribute.Name == RdfNamespace + "about") continue;
                    if (!IsDigitalSourceType(attribute.Name)) return false;
                    found = true;
                }
                foreach (XElement property in description.Elements()) {
                    if (!IsDigitalSourceType(property.Name) || property.HasElements) return false;
                    found = true;
                }
            }
        }
        return found;
    }

    private static bool IsDigitalSourceType(XName name) =>
        name.NamespaceName == IptcNamespace && name.LocalName == "DigitalSourceType";

    internal static OfficeProvenanceDigitalSourceKind Classify(string value) {
        if (value == null) return OfficeProvenanceDigitalSourceKind.Unknown;
        string normalized = value.Trim();
        if (normalized.StartsWith(VocabularyPrefix, StringComparison.Ordinal)) normalized = normalized.Substring(VocabularyPrefix.Length);
        return normalized switch {
            "digitalCapture" or "computationalCapture" => OfficeProvenanceDigitalSourceKind.DigitalCapture,
            "algorithmicMedia" => OfficeProvenanceDigitalSourceKind.AlgorithmicMedia,
            "trainedAlgorithmicMedia" => OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia,
            "compositeWithTrainedAlgorithmicMedia" or "compositeSynthetic" => OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia,
            "compositeCapture" => OfficeProvenanceDigitalSourceKind.CompositeCapture,
            "negativeFilm" or "positiveFilm" or "print" or
            "minorHumanEdits" or "humanEdits" or
            "softwareImage" or "digitalArt" or "digitalCreation" or
            "algorithmicallyEnhanced" or "dataDrivenMedia" or
            "virtualRecording" or "screenCapture" or "composite" => OfficeProvenanceDigitalSourceKind.Other,
            _ => OfficeProvenanceDigitalSourceKind.Unknown
        };
    }

    private static IEnumerable<XmpValue> FindValues(XDocument document, bool allowDirectRootIptc = false) {
        if (allowDirectRootIptc && document.Root != null) {
            foreach (XAttribute attribute in document.Root.Attributes()) {
                if (attribute.Name.NamespaceName == IptcNamespace && attribute.Name.LocalName == "DigitalSourceType") {
                    yield return new XmpValue(attribute.Value, Classify(attribute.Value), attribute, null, isStructurallyValid: true);
                }
            }
        }
        foreach (XElement description in document.Descendants(RdfNamespace + "Description")
            .Where(candidate => candidate.Ancestors().Any(ancestor => ancestor.Name == RdfNamespace + "RDF"))) {
            if (IsInsideXmlLiteral(description)) continue;
            foreach (XAttribute attribute in description.Attributes()) {
                if (attribute.Name.NamespaceName == IptcNamespace && attribute.Name.LocalName == "DigitalSourceType") {
                    yield return new XmpValue(attribute.Value, Classify(attribute.Value), attribute, null, isStructurallyValid: true);
                }
            }
            foreach (XElement element in description.Elements()) {
                if (IsInsideXmlLiteral(element)) continue;
                if (element.Name.NamespaceName == IptcNamespace && element.Name.LocalName == "DigitalSourceType") {
                    XAttribute? resource = element.Attribute(RdfNamespace + "resource");
                    string value = resource?.Value ?? element.Value;
                    yield return new XmpValue(value, Classify(value), resource, element, IsScalarRdfProperty(element, resource));
                }
            }
        }
    }

    private static bool IsInsideXmlLiteral(XElement element) =>
        element.AncestorsAndSelf().Any(candidate =>
            string.Equals(candidate.Attribute(RdfNamespace + "parseType")?.Value, "Literal", StringComparison.Ordinal));

    private static bool IsScalarRdfProperty(XElement element, XAttribute? resource) {
        if (element.HasElements || element.Nodes().Any(node => node is not XText)) return false;
        if (resource != null) {
            if (!string.IsNullOrWhiteSpace(element.Value)) return false;
            return element.Attributes().All(attribute => attribute.IsNamespaceDeclaration || attribute == resource);
        }
        if (element.Attribute(RdfNamespace + "parseType") != null ||
            element.Attribute(RdfNamespace + "nodeID") != null) return false;
        return element.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
            attribute.Name == XNamespace.Xml + "lang" || attribute.Name == RdfNamespace + "datatype");
    }

    private static bool TryLoad(byte[] packet, OfficeProvenanceOptions options, out XDocument? document) {
        return OfficeProvenanceXml.TryLoadDocument(packet, options, out document);
    }

    private sealed class XmpValue {
        internal XmpValue(string value, OfficeProvenanceDigitalSourceKind kind, XAttribute? attribute, XElement? element,
            bool isStructurallyValid) {
            Value = value;
            Kind = kind;
            Attribute = attribute;
            Element = element;
            IsStructurallyValid = isStructurallyValid;
        }
        internal string Value { get; }
        internal OfficeProvenanceDigitalSourceKind Kind { get; }
        internal XAttribute? Attribute { get; }
        internal XElement? Element { get; }
        internal bool IsStructurallyValid { get; }
    }
}

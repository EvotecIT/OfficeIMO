using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceXmp {
    private const string IptcNamespace = "http://iptc.org/std/Iptc4xmpExt/2008-02-29/";
    private static readonly XNamespace RdfNamespace = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
    private const string VocabularyPrefix = "http://cv.iptc.org/newscodes/digitalsourcetype/";

    internal static void Inspect(byte[] packet, OfficeProvenanceOptions options, OfficeProvenanceContext context, string location,
        bool carrierIsStructurallyValid = true) {
        if (!TryLoad(packet, options, out XDocument? document) || document == null) return;
        int index = 0;
        foreach (XmpValue value in FindValues(document)) {
            context.Add(new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.IptcDigitalSourceType,
                $"{location}/DigitalSourceType[{index++}]",
                isStructurallyValid: carrierIsStructurallyValid && value.Kind != OfficeProvenanceDigitalSourceKind.Unknown,
                value: value.Value,
                digitalSourceKind: value.Kind));
        }
    }

    internal static bool TryRemoveAiDeclarations(
        byte[] packet,
        OfficeProvenanceRemovalOptions options,
        string location,
        List<OfficeProvenanceChange> changes,
        out byte[] output) {
        output = packet;
        if (!options.RemoveAiSourceMetadata || !TryLoad(packet, options.Limits, out XDocument? document) || document == null) return false;
        XmpValue[] values = FindValues(document).ToArray();
        bool changed = false;
        int index = 0;
        foreach (XmpValue value in values) {
            bool remove = value.Kind == OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia ||
                value.Kind == OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia;
            if (remove) {
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
        using var stream = new MemoryStream();
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

    private static IEnumerable<XmpValue> FindValues(XDocument document) {
        foreach (XElement element in document.Descendants()) {
            if (!IsRdfPropertyContext(element)) continue;
            if (element.Name.NamespaceName == IptcNamespace && element.Name.LocalName == "DigitalSourceType") {
                XAttribute? resource = element.Attribute(RdfNamespace + "resource");
                string value = resource?.Value ?? element.Value;
                yield return new XmpValue(value, Classify(value), resource, element);
            }
            foreach (XAttribute attribute in element.Attributes()) {
                if (attribute.Name.NamespaceName == IptcNamespace && attribute.Name.LocalName == "DigitalSourceType") {
                    yield return new XmpValue(attribute.Value, Classify(attribute.Value), attribute, null);
                }
            }
        }
    }

    private static bool IsRdfPropertyContext(XElement element) {
        XElement? description = element.AncestorsAndSelf().FirstOrDefault(candidate => candidate.Name == RdfNamespace + "Description");
        return description != null && description.Ancestors().Any(candidate => candidate.Name == RdfNamespace + "RDF");
    }

    private static bool TryLoad(byte[] packet, OfficeProvenanceOptions options, out XDocument? document) {
        return OfficeProvenanceXml.TryLoadDocument(packet, options, out document);
    }

    private sealed class XmpValue {
        internal XmpValue(string value, OfficeProvenanceDigitalSourceKind kind, XAttribute? attribute, XElement? element) {
            Value = value;
            Kind = kind;
            Attribute = attribute;
            Element = element;
        }
        internal string Value { get; }
        internal OfficeProvenanceDigitalSourceKind Kind { get; }
        internal XAttribute? Attribute { get; }
        internal XElement? Element { get; }
    }
}

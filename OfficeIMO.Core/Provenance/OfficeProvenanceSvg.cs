using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Provenance;

internal static class OfficeProvenanceSvg {
    private static readonly XNamespace C2paNamespace = "http://c2pa.org/manifest";
    private static readonly XNamespace XmpNamespace = "adobe:ns:meta/";
    private static readonly XNamespace RdfNamespace = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
    private const string IptcNamespace = "http://iptc.org/std/Iptc4xmpExt/2008-02-29/";

    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        XDocument document = Load(data, options);
        IReadOnlyList<SvgCarrier> carriers = FindCarriers(document);
        int manifestCount = carriers.Count(carrier => carrier.Kind == SvgCarrierKind.Manifest);
        int xmpCarrierCount = carriers.Count(carrier => carrier.Kind == SvgCarrierKind.Xmp);
        int manifestIndex = 0;
        int xmpIndex = 0;
        foreach (SvgCarrier carrier in carriers) {
            if (carrier.Kind == SvgCarrierKind.Manifest) {
                XElement element = carrier.Element;
                string value = element.Value.Trim();
                byte[] manifest = Array.Empty<byte>();
                bool decoded = HasOnlyTextContent(element) && TryDecode(value, options.MaxManifestBytes, out manifest);
                bool valid = manifestCount == 1 && decoded && OfficeC2paManifestStore.IsValid(
                    manifest, 0, manifest.Length, options.MaxManifestBytes, options.MaxContainerEntries, out _);
                context.Add(new OfficeProvenanceEvidence(
                    OfficeProvenanceCarrierKind.C2paManifest,
                    $"SVG/metadata/c2pa:manifest[{manifestIndex++}]",
                    valid,
                    decoded ? manifest.Length : 0));
            } else {
                OfficeProvenanceXmp.Inspect(
                    SerializeElement(carrier.Element),
                    options,
                    context,
                    $"SVG/XMP[{xmpIndex++}]",
                    carrierIsStructurallyValid: xmpCarrierCount == 1);
            }
        }
    }

    internal static byte[] Remove(
        byte[] data,
        OfficeProvenanceRemovalOptions options,
        List<OfficeProvenanceChange> changes,
        out bool reserialized) {
        reserialized = false;
        if (!options.RemoveC2paManifests && !options.RemoveAiSourceMetadata) return (byte[])data.Clone();
        XDocument document = Load(data, options.Limits);
        IReadOnlyList<SvgCarrier> carriers = FindCarriers(document);
        int manifestCount = carriers.Count(carrier => carrier.Kind == SvgCarrierKind.Manifest);
        int xmpCarrierCount = carriers.Count(carrier => carrier.Kind == SvgCarrierKind.Xmp);
        int manifestIndex = 0;
        int xmpIndex = 0;
        foreach (SvgCarrier carrier in carriers) {
            if (carrier.Kind == SvgCarrierKind.Xmp) {
                string location = $"SVG/XMP[{xmpIndex++}]";
                if (!options.RemoveAiSourceMetadata || !OfficeProvenanceXmp.TryRemoveAiDeclarations(
                    SerializeElement(carrier.Element),
                    options,
                    location,
                    changes,
                    out byte[] cleanedXmp,
                    carrierIsStructurallyValid: xmpCarrierCount == 1)) continue;
                carrier.Element.ReplaceWith(LoadElement(cleanedXmp, options.Limits));
                reserialized = true;
                continue;
            }

            XElement element = carrier.Element;
            int index = manifestIndex++;
            byte[] manifest = Array.Empty<byte>();
            bool decoded = HasOnlyTextContent(element) && TryDecode(element.Value.Trim(), options.Limits.MaxManifestBytes, out manifest);
            bool valid = manifestCount == 1 && decoded && OfficeC2paManifestStore.IsValid(
                manifest, 0, manifest.Length, options.Limits.MaxManifestBytes, options.Limits.MaxContainerEntries, out _);
            if (!options.RemoveC2paManifests || !valid && options.RequireStructurallyValidCarrier) continue;
            element.Remove();
            changes.Add(new OfficeProvenanceChange(
                OfficeProvenanceCarrierKind.C2paManifest,
                $"SVG/metadata/c2pa:manifest[{index}]",
                0));
        }
        if (changes.Count == 0) return (byte[])data.Clone();
        using var output = new MemoryStream();
        var settings = new XmlWriterSettings {
            Encoding = new UTF8Encoding(false),
            Indent = false,
            OmitXmlDeclaration = document.Declaration == null,
            NewLineHandling = NewLineHandling.None
        };
        using (XmlWriter writer = XmlWriter.Create(output, settings)) document.Save(writer);
        reserialized = true;
        return output.ToArray();
    }

    private static XDocument Load(byte[] data, OfficeProvenanceOptions options) {
        ValidateMaterializedNodeBudget(data, options);
        XmlReaderSettings settings = CreateReaderSettings(options);
        using var stream = new MemoryStream(data, writable: false);
        using XmlReader reader = XmlReader.Create(stream, settings);
        XDocument document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        XElement? root = document.Root;
        if (root == null || root.Name.LocalName != "svg" || root.Name.NamespaceName != "http://www.w3.org/2000/svg") {
            throw new InvalidDataException("SVG root element is invalid.");
        }
        return document;
    }

    private static bool TryDecode(string value, long maximumBytes, out byte[] manifest) {
        manifest = Array.Empty<byte>();
        if (value.Length == 0) return false;
        long maximumEncodedBytes = maximumBytes > (long.MaxValue - 2L) / 4L * 3L
            ? long.MaxValue
            : ((maximumBytes + 2L) / 3L) * 4L;
        if (value.Length > maximumEncodedBytes) {
            throw OfficeProvenanceLimitException.Create("The SVG provenance manifest exceeds the configured manifest limit.");
        }
        try {
            manifest = Convert.FromBase64String(value);
            if (manifest.LongLength > maximumBytes) {
                throw OfficeProvenanceLimitException.Create("The SVG provenance manifest exceeds the configured manifest limit.");
            }
            return true;
        } catch (FormatException) {
            return false;
        }
    }

    private static bool IsManifestElement(XElement element) => element.Parent != null &&
        element.Parent.Name.LocalName == "metadata" &&
        element.Parent.Name.NamespaceName == "http://www.w3.org/2000/svg";

    private static bool HasOnlyTextContent(XElement element) => element.Nodes().All(node => node is XText);

    private static IReadOnlyList<SvgCarrier> FindCarriers(XDocument document) {
        var carriers = new List<SvgCarrier>();
        carriers.AddRange(document.Descendants(C2paNamespace + "manifest")
            .Where(IsManifestElement)
            .Select(static element => new SvgCarrier(element, SvgCarrierKind.Manifest)));
        carriers.AddRange(FindXmpRoots(document)
            .Select(static element => new SvgCarrier(element, SvgCarrierKind.Xmp)));
        carriers.Sort(static (left, right) => XNode.DocumentOrderComparer.Compare(left.Element, right.Element));
        return carriers;
    }

    private static IEnumerable<XElement> FindXmpRoots(XDocument document) {
        var roots = new List<XElement>();
        roots.AddRange(document.Descendants(XmpNamespace + "xmpmeta")
            .Where(element => !element.Ancestors(XmpNamespace + "xmpmeta").Any())
            .Where(element => element.Ancestors().Any(IsSvgMetadataElement)));
        XElement[] directIptcScopes = document.Descendants()
            .Where(ContainsDirectIptcDeclaration)
            .Where(element => !element.Ancestors().Any(ancestor => ancestor.Name == XmpNamespace + "xmpmeta"))
            .Where(element => element.Ancestors().Any(IsSvgMetadataElement))
            .Select(GetDirectIptcScope)
            .Distinct()
            .ToArray();
        var directIptcScopeSet = new HashSet<XElement>(directIptcScopes);
        roots.AddRange(directIptcScopes.Where(element =>
            !element.Ancestors().Any(directIptcScopeSet.Contains)));
        return roots;
    }

    private enum SvgCarrierKind {
        Manifest,
        Xmp
    }

    private readonly struct SvgCarrier {
        internal SvgCarrier(XElement element, SvgCarrierKind kind) {
            Element = element;
            Kind = kind;
        }

        internal XElement Element { get; }
        internal SvgCarrierKind Kind { get; }
    }

    private static bool ContainsDirectIptcDeclaration(XElement element) =>
        element.Name.NamespaceName == IptcNamespace ||
        element.Attributes().Any(attribute => attribute.Name.NamespaceName == IptcNamespace);

    private static XElement GetDirectIptcScope(XElement element) {
        XElement? rdf = element.AncestorsAndSelf().FirstOrDefault(ancestor => ancestor.Name == RdfNamespace + "RDF");
        if (rdf != null) return rdf;
        return element.Name.NamespaceName == IptcNamespace && element.Parent != null ? element.Parent : element;
    }

    private static bool IsSvgMetadataElement(XElement element) =>
        element.Name.LocalName == "metadata" &&
        element.Name.NamespaceName == "http://www.w3.org/2000/svg";

    private static byte[] SerializeElement(XElement element) {
        using var output = new MemoryStream();
        var settings = new XmlWriterSettings {
            Encoding = new UTF8Encoding(false),
            Indent = false,
            OmitXmlDeclaration = true,
            NewLineHandling = NewLineHandling.None
        };
        using (XmlWriter writer = XmlWriter.Create(output, settings)) element.Save(writer);
        return output.ToArray();
    }

    private static XElement LoadElement(byte[] data, OfficeProvenanceOptions options) {
        ValidateMaterializedNodeBudget(data, options);
        XmlReaderSettings settings = CreateReaderSettings(options);
        using var stream = new MemoryStream(data, writable: false);
        using XmlReader reader = XmlReader.Create(stream, settings);
        return XElement.Load(reader, LoadOptions.PreserveWhitespace);
    }

    private static XmlReaderSettings CreateReaderSettings(OfficeProvenanceOptions options) =>
        new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = options.MaxAssetBytes,
            MaxCharactersFromEntities = 0,
            IgnoreWhitespace = false
        };

    private static void ValidateMaterializedNodeBudget(byte[] data, OfficeProvenanceOptions options) {
        using var stream = new MemoryStream(data, writable: false);
        using XmlReader reader = XmlReader.Create(stream, CreateReaderSettings(options));
        int materializedNodes = 0;
        while (reader.Read()) {
            if (reader.Depth > 256) throw OfficeProvenanceLimitException.Create("SVG exceeds the configured XML depth limit.");
            switch (reader.NodeType) {
                case XmlNodeType.Element:
                    ReserveMaterializedNodes(ref materializedNodes, 1 + reader.AttributeCount, options.MaxContainerEntries);
                    break;
                case XmlNodeType.Text:
                case XmlNodeType.CDATA:
                case XmlNodeType.ProcessingInstruction:
                case XmlNodeType.Comment:
                case XmlNodeType.Whitespace:
                case XmlNodeType.SignificantWhitespace:
                    ReserveMaterializedNodes(ref materializedNodes, 1, options.MaxContainerEntries);
                    break;
            }
        }
    }

    private static void ReserveMaterializedNodes(ref int total, int count, int maximum) {
        if (count < 0 || total > maximum - count) {
            throw OfficeProvenanceLimitException.Create("SVG exceeds the configured XML node limit.");
        }
        total += count;
    }
}

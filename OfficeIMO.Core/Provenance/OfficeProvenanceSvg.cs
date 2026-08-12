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
    private const string IptcNamespace = "http://iptc.org/std/Iptc4xmpExt/2008-02-29/";

    internal static void Inspect(byte[] data, OfficeProvenanceOptions options, OfficeProvenanceContext context) {
        XDocument document = Load(data, options);
        int index = 0;
        foreach (XElement element in document.Descendants(C2paNamespace + "manifest").Where(IsManifestElement)) {
            string value = element.Value.Trim();
            bool decoded = TryDecode(value, options.MaxManifestBytes, out byte[] manifest);
            bool valid = decoded && OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, options.MaxManifestBytes, out _);
            context.Add(new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.C2paManifest,
                $"SVG/metadata/c2pa:manifest[{index++}]",
                valid,
                decoded ? manifest.Length : 0));
        }
        int xmpIndex = 0;
        foreach (XElement xmp in FindXmpRoots(document)) {
            OfficeProvenanceXmp.Inspect(SerializeElement(xmp), options, context, $"SVG/XMP[{xmpIndex++}]");
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
        XElement[] xmpRoots = FindXmpRoots(document).ToArray();
        for (int index = 0; index < xmpRoots.Length; index++) {
            XElement xmp = xmpRoots[index];
            if (!OfficeProvenanceXmp.TryRemoveAiDeclarations(
                SerializeElement(xmp),
                options,
                $"SVG/XMP[{index}]",
                changes,
                out byte[] cleanedXmp)) continue;
            xmp.ReplaceWith(LoadElement(cleanedXmp, options.Limits));
            reserialized = true;
        }
        XElement[] manifests = document.Descendants(C2paNamespace + "manifest").Where(IsManifestElement).ToArray();
        for (int index = 0; index < manifests.Length; index++) {
            XElement element = manifests[index];
            bool decoded = TryDecode(element.Value.Trim(), options.Limits.MaxManifestBytes, out byte[] manifest);
            bool valid = decoded && OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, options.Limits.MaxManifestBytes, out _);
            if (!options.RemoveC2paManifests) continue;
            if (!valid && options.RequireStructurallyValidCarrier) continue;
            string location = $"SVG/metadata/c2pa:manifest[{index}]";
            element.Remove();
            changes.Add(new OfficeProvenanceChange(OfficeProvenanceCarrierKind.C2paManifest, location, 0));
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
        if (value.Length == 0 || value.Length > maximumBytes * 2L) return false;
        try {
            manifest = Convert.FromBase64String(value);
            return manifest.LongLength <= maximumBytes;
        } catch (FormatException) {
            return false;
        }
    }

    private static bool IsManifestElement(XElement element) => element.Parent != null &&
        element.Parent.Name.LocalName == "metadata" &&
        element.Parent.Name.NamespaceName == "http://www.w3.org/2000/svg";

    private static IEnumerable<XElement> FindXmpRoots(XDocument document) {
        var roots = new List<XElement>();
        roots.AddRange(document.Descendants(XmpNamespace + "xmpmeta"));
        XElement[] directIptcScopes = document.Descendants()
            .Where(ContainsDirectIptcDeclaration)
            .Where(element => !element.Ancestors().Any(ancestor => ancestor.Name == XmpNamespace + "xmpmeta"))
            .Where(element => element.Ancestors().Any(IsSvgMetadataElement))
            .Select(GetDirectIptcScope)
            .Distinct()
            .ToArray();
        roots.AddRange(directIptcScopes.Where(element =>
            !element.Ancestors().Any(ancestor => directIptcScopes.Contains(ancestor))));
        return roots;
    }

    private static bool ContainsDirectIptcDeclaration(XElement element) =>
        element.Name.NamespaceName == IptcNamespace ||
        element.Attributes().Any(attribute => attribute.Name.NamespaceName == IptcNamespace);

    private static XElement GetDirectIptcScope(XElement element) =>
        element.Name.NamespaceName == IptcNamespace && element.Parent != null
            ? element.Parent
            : element;

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
            if (reader.Depth > 256) throw new InvalidDataException("SVG exceeds the configured XML depth limit.");
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
            throw new InvalidDataException("SVG exceeds the configured XML node limit.");
        }
        total += count;
    }
}

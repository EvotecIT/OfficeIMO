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
        OfficeProvenanceXmp.Inspect(data, options, context, "SVG/XMP");
    }

    internal static byte[] Remove(
        byte[] data,
        OfficeProvenanceRemovalOptions options,
        List<OfficeProvenanceChange> changes,
        out bool reserialized) {
        reserialized = false;
        if (!options.RemoveC2paManifests && !options.RemoveAiSourceMetadata) return (byte[])data.Clone();
        byte[] working = data;
        if (OfficeProvenanceXmp.TryRemoveAiDeclarations(data, options, "SVG/XMP", changes, out byte[] cleanedXmp)) {
            working = cleanedXmp;
            reserialized = true;
        }
        XDocument document = Load(working, options.Limits);
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
        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = options.MaxAssetBytes,
            MaxCharactersFromEntities = 0,
            IgnoreWhitespace = false
        };
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
}

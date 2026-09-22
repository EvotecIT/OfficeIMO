using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
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
#if NET8_0_OR_GREATER
        if (TryInspectManifestOnly(data, options, context)) return;
#endif
        XDocument document = Load(data, options);
        IReadOnlyList<SvgCarrier> carriers = FindCarriers(document);
        int manifestCount = carriers.Count(carrier => carrier.Kind == SvgCarrierKind.Manifest);
        int xmpCarrierCount = CountLogicalXmpCarriers(carriers);
        int manifestIndex = 0;
        int xmpIndex = 0;
        foreach (SvgCarrier carrier in carriers) {
            if (carrier.Kind == SvgCarrierKind.Manifest) {
                XElement element = carrier.Element;
                string value = element.Value.Trim();
                int manifestLength = 0;
                bool structurallyValid = false;
                bool decoded = HasOnlyTextContent(element) && TryValidateManifest(
                    value, options.MaxManifestBytes, options.MaxContainerEntries, out manifestLength, out structurallyValid);
                bool valid = manifestCount == 1 && decoded && structurallyValid;
                context.Add(new OfficeProvenanceEvidence(
                    OfficeProvenanceCarrierKind.C2paManifest,
                    $"SVG/metadata/c2pa:manifest[{manifestIndex++}]",
                    valid,
                    decoded ? manifestLength : 0));
            } else {
                OfficeProvenanceXmp.Inspect(
                    SerializeElement(carrier.Element),
                    options,
                    context,
                    $"SVG/XMP[{xmpIndex++}]",
                    carrierIsStructurallyValid: xmpCarrierCount == 1,
                    allowDirectRootIptc: IsSvgMetadataElement(carrier.Element));
            }
        }
    }

    internal static byte[] Remove(
        byte[] data,
        OfficeProvenanceRemovalOptions options,
        OfficeProvenanceReport before,
        List<OfficeProvenanceChange> changes,
        out bool reserialized) {
        reserialized = false;
        if (!options.RemoveC2paManifests && !options.RemoveAiSourceMetadata) {
            return OfficeProvenanceBinary.CloneForOutput(data, options.EffectiveMaxOutputBytes);
        }
        XDocument document = Load(data, options.Limits);
        IReadOnlyList<SvgCarrier> carriers = FindCarriers(document);
        int xmpCarrierCount = CountLogicalXmpCarriers(carriers);
        var manifestIndexes = new Dictionary<XElement, int>();
        var xmpIndexes = new Dictionary<XElement, int>();
        int manifestIndex = 0;
        int xmpIndex = 0;
        foreach (SvgCarrier carrier in carriers) {
            if (carrier.Kind == SvgCarrierKind.Manifest) manifestIndexes[carrier.Element] = manifestIndex++;
            else xmpIndexes[carrier.Element] = xmpIndex++;
        }
        // Process descendants first so replacing an XMP scope cannot detach a carrier
        // that was discovered within that scope before it is removed. Restore the public
        // change list to source order after all mutations have completed.
        int initialChangeCount = changes.Count;
        var orderedChanges = new List<(int Order, OfficeProvenanceChange Change)>();
        for (int carrierOrder = carriers.Count - 1; carrierOrder >= 0; carrierOrder--) {
            SvgCarrier carrier = carriers[carrierOrder];
            int previousChangeCount = changes.Count;
            if (carrier.Kind == SvgCarrierKind.Xmp) {
                string location = $"SVG/XMP[{xmpIndexes[carrier.Element]}]";
                if (options.RemoveAiSourceMetadata && OfficeProvenanceXmp.TryRemoveAiDeclarations(
                    SerializeElement(carrier.Element),
                    options,
                    location,
                    changes,
                    out byte[] cleanedXmp,
                    carrierIsStructurallyValid: xmpCarrierCount == 1,
                    allowDirectRootIptc: IsSvgMetadataElement(carrier.Element))) {
                    carrier.Element.ReplaceWith(LoadElement(cleanedXmp, options.Limits));
                    reserialized = true;
                }
            } else {
                int index = manifestIndexes[carrier.Element];
                XElement element = carrier.Element;
                string manifestLocation = $"SVG/metadata/c2pa:manifest[{index}]";
                bool valid = before.Evidence.Any(item =>
                    item.Carrier == OfficeProvenanceCarrierKind.C2paManifest &&
                    item.Location == manifestLocation &&
                    item.IsStructurallyValid);
                if (options.RemoveC2paManifests && (valid || !options.RequireStructurallyValidCarrier)) {
                    element.Remove();
                    changes.Add(new OfficeProvenanceChange(
                        OfficeProvenanceCarrierKind.C2paManifest,
                        manifestLocation,
                        0));
                }
            }
            for (int changeIndex = previousChangeCount; changeIndex < changes.Count; changeIndex++) {
                orderedChanges.Add((carrierOrder, changes[changeIndex]));
            }
        }
        changes.RemoveRange(initialChangeCount, changes.Count - initialChangeCount);
        changes.AddRange(orderedChanges.OrderBy(item => item.Order).Select(item => item.Change));
        if (changes.Count == 0) return OfficeProvenanceBinary.CloneForOutput(data, options.EffectiveMaxOutputBytes);
        using var output = new OfficeProvenanceBoundedMemoryStream(options.EffectiveMaxOutputBytes);
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

    private static bool TryValidateManifest(
        string value,
        long maximumBytes,
        int maximumEntries,
        out int manifestLength,
        out bool structurallyValid) {
#if NET8_0_OR_GREATER
        return TryValidateManifest(value.AsSpan(), maximumBytes, maximumEntries, out manifestLength, out structurallyValid);
#else
        manifestLength = 0;
        structurallyValid = false;
        if (value.Length == 0) return false;
        long maximumEncodedBytes = maximumBytes > (long.MaxValue - 2L) / 4L * 3L
            ? long.MaxValue
            : ((maximumBytes + 2L) / 3L) * 4L;
        if (value.Length > maximumEncodedBytes) {
            throw OfficeProvenanceLimitException.Create("The SVG provenance manifest exceeds the configured manifest limit.");
        }
        try {
            byte[] manifest = Convert.FromBase64String(value);
            manifestLength = manifest.Length;
            if (manifest.LongLength > maximumBytes) {
                throw OfficeProvenanceLimitException.Create("The SVG provenance manifest exceeds the configured manifest limit.");
            }
            structurallyValid = OfficeC2paManifestStore.IsValid(
                manifest, 0, manifest.Length, maximumBytes, maximumEntries, out _);
            return true;
        } catch (FormatException) {
            return false;
        }
#endif
    }

#if NET8_0_OR_GREATER
    private static bool TryValidateManifest(
        ReadOnlySpan<char> value,
        long maximumBytes,
        int maximumEntries,
        out int manifestLength,
        out bool structurallyValid) {
        manifestLength = 0;
        structurallyValid = false;
        if (value.Length == 0) return false;
        long maximumEncodedBytes = maximumBytes > (long.MaxValue - 2L) / 4L * 3L
            ? long.MaxValue
            : ((maximumBytes + 2L) / 3L) * 4L;
        if (value.Length > maximumEncodedBytes) {
            throw OfficeProvenanceLimitException.Create("The SVG provenance manifest exceeds the configured manifest limit.");
        }
        int maximumDecodedLength = checked((value.Length / 4) * 3 + 3);
        byte[] manifest = ArrayPool<byte>.Shared.Rent(maximumDecodedLength);
        try {
            if (!Convert.TryFromBase64Chars(value, manifest, out manifestLength)) return false;
            if (manifestLength > maximumBytes) {
                throw OfficeProvenanceLimitException.Create("The SVG provenance manifest exceeds the configured manifest limit.");
            }
            structurallyValid = OfficeC2paManifestStore.IsValid(
                manifest, 0, manifestLength, maximumBytes, maximumEntries, out _);
            return true;
        } finally {
            ArrayPool<byte>.Shared.Return(manifest, clearArray: true);
        }
    }
#endif

#if NET8_0_OR_GREATER
    private static bool TryInspectManifestOnly(
        byte[] data,
        OfficeProvenanceOptions options,
        OfficeProvenanceContext context) {
        using var stream = new MemoryStream(data, writable: false);
        using XmlReader reader = XmlReader.Create(stream, CreateReaderSettings(options));
        var metadataDepths = new Stack<int>();
        var manifests = new List<SvgManifestEvidence>();
        int materializedNodes = 0;
        bool rootSeen = false;
        while (reader.Read()) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType == XmlNodeType.Element && reader.Depth > 256) throw OfficeProvenanceLimitException.Create("SVG exceeds the configured XML depth limit.");
            if (reader.NodeType == XmlNodeType.Element) {
                ReserveMaterializedNodes(ref materializedNodes, 1 + reader.AttributeCount, options.MaxContainerEntries);
                if (!rootSeen) {
                    rootSeen = true;
                    if (reader.Depth != 0 || reader.LocalName != "svg" || reader.NamespaceURI != "http://www.w3.org/2000/svg") {
                        throw new InvalidDataException("SVG root element is invalid.");
                    }
                }
                bool insideMetadata = metadataDepths.Count != 0 && reader.Depth > metadataDepths.Peek();
                if (reader.LocalName == "metadata" && reader.NamespaceURI == "http://www.w3.org/2000/svg" &&
                    IsXmpCarrier(reader)) return false;
                if (insideMetadata && IsXmpCarrier(reader)) return false;
                if (insideMetadata && reader.Depth == metadataDepths.Peek() + 1 &&
                    reader.LocalName == "manifest" && reader.NamespaceURI == C2paNamespace.NamespaceName) {
                    if (manifests.Count >= options.MaxCarriers) {
                        throw OfficeProvenanceLimitException.Create("SVG exceeds the configured carrier limit.");
                    }
                    SvgManifestEvidence manifest = ReadManifestElement(reader, data.Length, options, ref materializedNodes);
                    if (manifest.RequiresFallback) return false;
                    manifests.Add(manifest);
                    continue;
                }
                if (reader.LocalName == "metadata" && reader.NamespaceURI == "http://www.w3.org/2000/svg" && !reader.IsEmptyElement) {
                    metadataDepths.Push(reader.Depth);
                }
            } else if (reader.NodeType == XmlNodeType.EndElement) {
                if (metadataDepths.Count != 0 && metadataDepths.Peek() == reader.Depth) metadataDepths.Pop();
            } else if (IsMaterializedTextNode(reader.NodeType)) {
                ReserveMaterializedNodes(ref materializedNodes, 1, options.MaxContainerEntries);
            }
        }
        if (!rootSeen) throw new InvalidDataException("SVG root element is invalid.");
        for (int index = 0; index < manifests.Count; index++) {
            SvgManifestEvidence manifest = manifests[index];
            context.Add(new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.C2paManifest,
                $"SVG/metadata/c2pa:manifest[{index}]",
                manifests.Count == 1 && manifest.Decoded && manifest.StructurallyValid,
                manifest.Decoded ? manifest.ManifestLength : 0));
        }
        return true;
    }

    private static SvgManifestEvidence ReadManifestElement(
        XmlReader reader,
        int encodedAssetLength,
        OfficeProvenanceOptions options,
        ref int materializedNodes) {
        if (reader.IsEmptyElement) return default;
        int manifestDepth = reader.Depth;
        long maximumBase64Chars = ((options.MaxManifestBytes + 2L) / 3L) * 4L;
        int maximumValueChars = (int)Math.Min(encodedAssetLength, maximumBase64Chars + 4096L);
        char[] value = ArrayPool<char>.Shared.Rent(Math.Min(4096, Math.Max(1, maximumValueChars)));
        char[] chunkBuffer = ArrayPool<char>.Shared.Rent(4096);
        int valueLength = 0;
        bool onlyText = true;
        try {
            using (XmlReader subtree = reader.ReadSubtree()) {
                bool first = true;
                while (subtree.Read()) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    if (first) {
                        first = false;
                        continue;
                    }
                    if (subtree.NodeType == XmlNodeType.Element && manifestDepth + subtree.Depth > 256) throw OfficeProvenanceLimitException.Create("SVG exceeds the configured XML depth limit.");
                    if (subtree.NodeType == XmlNodeType.Element) {
                        ReserveMaterializedNodes(ref materializedNodes, 1 + subtree.AttributeCount, options.MaxContainerEntries);
                        onlyText = false;
                        if (IsXmpCarrier(subtree) || IsManifestElement(subtree)) {
                            return new SvgManifestEvidence(false, false, 0, requiresFallback: true);
                        }
                    } else if (IsMaterializedTextNode(subtree.NodeType)) {
                        ReserveMaterializedNodes(ref materializedNodes, 1, options.MaxContainerEntries);
                        if (subtree.NodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace) {
                            if (!subtree.CanReadValueChunk) {
                                throw OfficeProvenanceLimitException.Create("SVG provenance manifest text cannot be read within the configured limit.");
                            }
                            int chunkLength;
                            while ((chunkLength = subtree.ReadValueChunk(chunkBuffer, 0, chunkBuffer.Length)) > 0) {
                                options.CancellationToken.ThrowIfCancellationRequested();
                                long required = (long)valueLength + chunkLength;
                                if (required > maximumValueChars) {
                                    throw OfficeProvenanceLimitException.Create("The SVG provenance manifest exceeds the configured manifest limit.");
                                }
                                if (required > value.Length) {
                                    int capacity = (int)Math.Min(maximumValueChars, Math.Max(required, (long)value.Length * 2L));
                                    char[] expanded = ArrayPool<char>.Shared.Rent(capacity);
                                    Array.Copy(value, expanded, valueLength);
                                    ArrayPool<char>.Shared.Return(value, clearArray: true);
                                    value = expanded;
                                }
                                Array.Copy(chunkBuffer, 0, value, valueLength, chunkLength);
                                valueLength += chunkLength;
                            }
                        } else {
                            onlyText = false;
                        }
                    }
                }
            }
            if (reader.Depth != manifestDepth || reader.NodeType != XmlNodeType.EndElement) {
                throw new InvalidDataException("SVG provenance manifest parsing did not end at its carrier boundary.");
            }
            if (!onlyText) return default;
            int start = 0;
            while (start < valueLength && char.IsWhiteSpace(value[start])) start++;
            while (valueLength > start && char.IsWhiteSpace(value[valueLength - 1])) valueLength--;
            bool decoded = TryValidateManifest(
                new ReadOnlySpan<char>(value, start, valueLength - start),
                options.MaxManifestBytes,
                options.MaxContainerEntries,
                out int manifestLength,
                out bool structurallyValid);
            return new SvgManifestEvidence(decoded, structurallyValid, manifestLength);
        } finally {
            ArrayPool<char>.Shared.Return(chunkBuffer, clearArray: true);
            ArrayPool<char>.Shared.Return(value, clearArray: true);
        }
    }

    private static bool IsXmpCarrier(XmlReader reader) {
        if (reader.NamespaceURI == XmpNamespace.NamespaceName && reader.LocalName == "xmpmeta") return true;
        if (reader.NamespaceURI == IptcNamespace) return true;
        if (!reader.HasAttributes) return false;
        bool hasIptcAttribute = false;
        if (reader.MoveToFirstAttribute()) {
            do {
                if (reader.NamespaceURI == IptcNamespace) {
                    hasIptcAttribute = true;
                    break;
                }
            } while (reader.MoveToNextAttribute());
            reader.MoveToElement();
        }
        return hasIptcAttribute;
    }

    private static bool IsManifestElement(XmlReader reader) =>
        reader.NamespaceURI == C2paNamespace.NamespaceName && reader.LocalName == "manifest";

    private static bool IsMaterializedTextNode(XmlNodeType nodeType) => nodeType is
        XmlNodeType.Text or
        XmlNodeType.CDATA or
        XmlNodeType.ProcessingInstruction or
        XmlNodeType.Comment or
        XmlNodeType.Whitespace or
        XmlNodeType.SignificantWhitespace;
#endif

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

    private static int CountLogicalXmpCarriers(IReadOnlyList<SvgCarrier> carriers) {
        int count = 0;
        foreach (SvgCarrier carrier in carriers) {
            if (carrier.Kind != SvgCarrierKind.Xmp) continue;
            var scopes = new HashSet<XElement>();
            foreach (XElement xmp in carrier.Element.DescendantsAndSelf(XmpNamespace + "xmpmeta")
                .Where(element => !element.Ancestors(XmpNamespace + "xmpmeta").Any())) {
                scopes.Add(xmp);
            }
            foreach (XElement declaration in carrier.Element.DescendantsAndSelf()
                .Where(ContainsDirectIptcDeclaration)
                .Where(element => !element.Ancestors(XmpNamespace + "xmpmeta").Any())) {
                scopes.Add(GetDirectIptcScope(declaration));
            }
            // A direct declaration on metadata belongs to its one nested XMP/RDF
            // scope, but multiple sibling scopes remain separate carriers.
            if (scopes.Count > 1) scopes.Remove(carrier.Element);
            count = checked(count + Math.Max(1, scopes.Count));
        }
        return count;
    }

    private static IEnumerable<XElement> FindXmpRoots(XDocument document) {
        var roots = new List<XElement>();
        roots.AddRange(document.Descendants(XmpNamespace + "xmpmeta")
            .Where(element => !element.Ancestors(XmpNamespace + "xmpmeta").Any())
            .Where(element => element.Ancestors().Any(IsSvgMetadataElement)));
        XElement[] directIptcScopes = document.Descendants()
            .Where(ContainsDirectIptcDeclaration)
            .Where(element => !element.Ancestors().Any(ancestor => ancestor.Name == XmpNamespace + "xmpmeta"))
            .Where(element => element.AncestorsAndSelf().Any(IsSvgMetadataElement))
            .Select(GetDirectIptcScope)
            .Distinct()
            .ToArray();
        var directIptcScopeSet = new HashSet<XElement>(directIptcScopes);
        roots.AddRange(directIptcScopes.Where(element =>
            !element.Ancestors().Any(directIptcScopeSet.Contains)));
        var rootSet = new HashSet<XElement>(roots);
        return roots.Distinct().Where(element => !element.Ancestors().Any(rootSet.Contains));
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

#if NET8_0_OR_GREATER
    private readonly struct SvgManifestEvidence {
        internal SvgManifestEvidence(bool decoded, bool structurallyValid, int manifestLength, bool requiresFallback = false) {
            Decoded = decoded;
            StructurallyValid = structurallyValid;
            ManifestLength = manifestLength;
            RequiresFallback = requiresFallback;
        }

        internal bool Decoded { get; }
        internal bool StructurallyValid { get; }
        internal int ManifestLength { get; }
        internal bool RequiresFallback { get; }
    }
#endif

    private static bool ContainsDirectIptcDeclaration(XElement element) =>
        element.Name.NamespaceName == IptcNamespace ||
        element.Attributes().Any(attribute => attribute.Name.NamespaceName == IptcNamespace);

    private static XElement GetDirectIptcScope(XElement element) {
        XElement? rdf = element.AncestorsAndSelf().FirstOrDefault(ancestor => ancestor.Name == RdfNamespace + "RDF");
        if (rdf != null) return rdf;
        if (element.Name == C2paNamespace + "manifest") {
            return element.Ancestors().FirstOrDefault(IsSvgMetadataElement) ?? element;
        }
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
            if (reader.NodeType == XmlNodeType.Element && reader.Depth > 256) throw OfficeProvenanceLimitException.Create("SVG exceeds the configured XML depth limit.");
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

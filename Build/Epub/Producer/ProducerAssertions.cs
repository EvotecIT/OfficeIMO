using OfficeIMO.Epub;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

internal static class ProducerAssertions {
    // Call only after EpubPublication.Load has enforced the input's archive/resource limits.
    internal static Dictionary<string, byte[]> Payloads(byte[] bytes) {
        using var zip = new ZipArchive(new MemoryStream(bytes));
        return zip.Entries.Where(entry => !entry.FullName.EndsWith('/')).ToDictionary(entry => entry.FullName, entry => {
            using var input = entry.Open(); using var output = new MemoryStream(); input.CopyTo(output); return output.ToArray();
        }, StringComparer.Ordinal);
    }

    internal static object Verify(EpubPublication source, EpubPublication edited, IReadOnlyDictionary<string, byte[]> original,
        IReadOnlyDictionary<string, byte[]> proposed, IReadOnlyDictionary<string, string> moves) {
        int assets = 0, overlays = 0, bodies = 0;
        foreach (var item in source.Manifest) {
            if (item.Reference.Kind != EpubReferenceKind.Container) continue;
            string oldPath = item.Reference.ContainerPath!;
            string path = moves.TryGetValue(oldPath, out string? moved) ? moved : oldPath;
            var current = edited.Manifest.Single(candidate => candidate.Id == item.Id);
            if (current.Reference.ContainerPath != path || current.MediaType != item.MediaType || current.MediaOverlayId != item.MediaOverlayId)
                throw new InvalidDataException("Resource identity, type or narration association changed: " + item.Id);
            if (item.MediaType == "application/xhtml+xml" && !(item.Properties ?? "").Split(' ').Contains("nav")) {
                XNamespace html = "http://www.w3.org/1999/xhtml";
                string originalText = source.GetContentXml(item.Id).Root!.Element(html + "body")!.Value;
                string currentText = edited.GetContentXml(item.Id).Root!.Element(html + "body")!.Value;
                if (item.Id == "xintroduction_001" && edited.Manifest.Any(resource => resource.Id == "editorial-second"))
                    currentText += edited.GetContentXml("editorial-second").Root!.Element(html + "body")!.Value;
                if (originalText != currentText) throw new InvalidDataException("Chapter body text or order changed: " + item.Id);
                bodies++;
            }
            bool xml = item.MediaType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase) || item.MediaType == "application/xml" || item.MediaType == "text/xml";
            if (!xml && item.MediaType != "text/css") {
                if (!original[oldPath].SequenceEqual(proposed[path])) throw new InvalidDataException("Asset bytes changed: " + oldPath);
                assets++;
            }
            if (item.MediaType == "application/smil+xml") {
                if (!XNode.DeepEquals(CanonicalOverlay(original[oldPath], oldPath, moves),
                    CanonicalOverlay(proposed[path], path, new Dictionary<string, string>())))
                    throw new InvalidDataException("SMIL content, timing or target identity changed: " + item.Id);
                overlays++;
            }
        }
        var before = Navigation(source.Read().TableOfContents, moves).ToArray();
        var after = Navigation(edited.Read().TableOfContents, new Dictionary<string, string>()).ToArray();
        int next = 0;
        foreach (var expected in before) {
            while (next < after.Length && after[next] != expected) next++;
            if (next == after.Length) throw new InvalidDataException("An original TOC label, target, order or nesting changed: " + expected);
            next++;
        }
        var changed = original.Where(pair => {
            string target = moves.TryGetValue(pair.Key, out string? moved) ? moved : pair.Key;
            return !proposed.TryGetValue(target, out byte[]? value) || !pair.Value.SequenceEqual(value);
        }).Select(pair => pair.Key).ToArray();
        return new { preservedChapterBodies = bodies, byteIdenticalNonXmlAssets = assets, structureTimingAndTargetsPreservedOverlays = overlays,
            preservedOriginalTocEntries = before.Length, resultingTocEntries = after.Length, changedSourcePayloads = changed };
    }

    private static XElement CanonicalOverlay(byte[] bytes, string owner, IReadOnlyDictionary<string, string> moves) {
        using var reader = XmlReader.Create(new MemoryStream(bytes), new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 64L * 1024 * 1024
        });
        var root = XDocument.Load(reader, LoadOptions.PreserveWhitespace).Root!;
        foreach (var attribute in root.DescendantsAndSelf().Attributes().Where(attribute =>
            attribute.Name == "src" || attribute.Name == XName.Get("textref", "http://www.idpf.org/2007/ops"))) {
            var reference = EpubReference.Resolve(owner, null, attribute.Value);
            if (reference.Kind != EpubReferenceKind.Container) throw new InvalidDataException("Unexpected non-container narration target.");
            string target = moves.TryGetValue(reference.ContainerPath!, out string? moved) ? moved : reference.ContainerPath!;
            attribute.Value = target + (reference.Query == null ? "" : "?" + reference.Query) +
                (reference.Fragment == null ? "" : "#" + reference.Fragment);
        }
        return root;
    }

    private static IEnumerable<(string Label, string? Target, string? Fragment, int Depth)> Navigation(
        IReadOnlyList<EpubNavigationItem> items, IReadOnlyDictionary<string, string> moves, int depth = 0) {
        foreach (var item in items) {
            yield return (item.Label, item.Target != null && moves.TryGetValue(item.Target, out string? moved) ? moved : item.Target, item.Fragment, depth);
            foreach (var child in Navigation(item.Children, moves, depth + 1)) yield return child;
        }
    }
}

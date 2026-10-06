using System.Threading;

namespace OfficeIMO.Epub;

public static partial class EpubManuscript {
    /// <summary>Plans heading boundaries before cloning so document-local relationships retain their original elements.</summary>
    private static HashSet<XElement> PlanChapterBoundaries(XElement body, int headingLevel,
        List<OfficeConversionFidelityDiagnostic> diagnostics, CancellationToken token) {
        var allowed = new HashSet<XElement>();
        if (headingLevel == 0) return allowed;
        var boundaries = new List<XElement>();
        var spans = new Dictionary<XElement, (int First, int Last)>();
        var ids = new Dictionary<string, XElement>(StringComparer.Ordinal);
        var maps = new Dictionary<string, XElement>(StringComparer.Ordinal);
        int chapter = 0;
        bool readable = false;
        Visit(body, true);

        // Range additions suppress each crossed boundary once, including overlapping/transitive relationships.
        // Element spans include descendants: a referenced container must retain its complete content.
        var crossings = new int[boundaries.Count + 2];
        foreach (var pair in spans) {
            token.ThrowIfCancellationRequested();
            foreach (XAttribute attribute in EpubContentIdentifiers.GetReferenceAttributes(pair.Key))
                foreach (string id in EpubContentIdentifiers.GetReferenceIds(attribute)) {
                    token.ThrowIfCancellationRequested();
                    if (ids.TryGetValue(id, out XElement? target)) KeepTogether(pair.Value, spans[target]);
                }
            string? map = EpubContentIdentifiers.GetImageMapReference(pair.Key);
            if (map != null && maps.TryGetValue(map, out XElement? imageMap)) KeepTogether(pair.Value, spans[imageMap]);
        }
        int active = 0;
        for (int index = 0; index < boundaries.Count; index++) {
            token.ThrowIfCancellationRequested();
            active += crossings[index + 1];
            if (active == 0) continue;
            allowed.Remove(boundaries[index]);
            AddDiagnostic(diagnostics, "EPUB_IMPORT_CHAPTER_BOUNDARY_PRESERVED",
                "A proposed chapter boundary was kept inside its chapter to preserve document-local relationships; the heading remains in navigation.",
                (string?)boundaries[index].Attribute("id") ?? boundaries[index].Value.Trim(), OfficeConversionLossKind.None);
        }
        return allowed;

        void KeepTogether((int First, int Last) source, (int First, int Last) target) {
            int first = Math.Min(source.First, target.First);
            int last = Math.Max(source.Last, target.Last);
            if (first == last) return;
            crossings[first + 1]++;
            crossings[last + 1]--;
        }

        void Visit(XElement element, bool canSplit) {
            token.ThrowIfCancellationRequested();
            int level = HeadingLevel(element);
            if (canSplit && level != 0 && level <= headingLevel) {
                if (readable) {
                    allowed.Add(element);
                    chapter++;
                    boundaries.Add(element);
                    readable = false;
                }
            }
            int first = chapter;
            foreach (XAttribute id in element.Attributes().Where(attribute => attribute.Name == "id" || attribute.Name == XNamespace.Xml + "id"))
                if (!ids.ContainsKey(id.Value)) ids.Add(id.Value, element);
            if (element.Name == Xhtml + "map" && (string?)element.Attribute("name") is string name && !maps.ContainsKey(name)) maps.Add(name, element);
            if (IsReadableMedia(element)) readable = true;
            bool childCanSplit = canSplit && (element == body || IsChapterContainer(element));
            foreach (XNode node in element.Nodes()) {
                token.ThrowIfCancellationRequested();
                if (node is XElement child) Visit(child, childCanSplit);
                else if (node is XText text && !string.IsNullOrWhiteSpace(text.Value)) readable = true;
            }
            spans.Add(element, (first, chapter));
        }
    }
}

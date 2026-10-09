namespace OfficeIMO.Xps;

public sealed partial class XpsDocument {
    /// <summary>The native DocumentStructure/StoryFragments namespace for this document's dialect.</summary>
    public XNamespace StructureNamespace => XpsPackage.Namespace(Format) + "/documentstructure";

    // Read only relationship-owned structure parts. An unrelated XML resource with
    // a similar element name must not acquire document semantics during an edit.
    internal IEnumerable<(string Part, XElement Markup)> ReadDocumentStructures(CancellationToken token = default, bool activeOnly = false) {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var document in activeOnly ? _documents.Distinct() : _documentCache.Values.AsEnumerable()) {
            token.ThrowIfCancellationRequested();
            var structure = ReadOwnedDocumentStructure(document, token);
            if (structure.HasValue && seen.Add(structure.Value.Part)) yield return structure.Value;
        }
    }

    private Dictionary<string, byte[]> PreserveDocumentStructure(StructureIndex next) {
        var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        int[] pageMap = MapPageOccurrences(next.Pages);
        foreach (var structure in ReadDocumentStructures()) {
            bool changed = false;
            // These are payload-global, one-based page numbers, not document-local
            // page indexes (ECMA-388 section 16.1.1.6).
            foreach (var story in structure.Markup.Elements(StructureNamespace + "Story").ToArray()) {
                bool removed = false;
                foreach (var reference in story.Elements(StructureNamespace + "StoryFragmentReference").ToArray()) {
                    string? value = (string?)reference.Attribute("Page");
                    if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int number) || number < 1 || number > _pages.Count)
                        throw new InvalidDataException("Unresolved story-fragment page number.");
                    int current = pageMap[number - 1];
                    if (current < 0) { reference.Remove(); changed = removed = true; }
                    else if (current != number - 1) { reference.SetAttributeValue("Page", current + 1); changed = true; }
                }
                if (removed && !story.Elements().Any()) story.Remove();
            }
            foreach (var outline in structure.Markup.Elements(StructureNamespace + "DocumentStructure.Outline")
                .Elements(StructureNamespace + "DocumentOutline").Elements(StructureNamespace + "OutlineEntry")) {
                string? uri = (string?)outline.Attribute("OutlineTarget");
                if (uri == null || uri.IndexOf(':') >= 0 || uri.StartsWith("#", StringComparison.Ordinal)) continue;
                string[] pieces = uri.Split('#');
                if (pieces.Length > 2) continue;
                string name;
                try { name = XpsPackage.Resolve(structure.Part, pieces[0]); } catch (InvalidDataException) { continue; }
                string? anchor = pieces.Length == 2 ? Uri.UnescapeDataString(pieces[1]) : null;
                int previous = LinkTargetPage(name, anchor);
                if (previous < 0) continue;
                var target = _pages[previous];
                int current = LinkTargetPage(next, name, anchor);
                if (current >= 0 && next.Pages[current] == target) continue;
                string fragment = anchor != null && target.HasNamedTarget(anchor) ? "#" + Uri.EscapeDataString(anchor) : "";
                outline.SetAttributeValue("OutlineTarget", "/" + target.PartName + fragment);
                changed = true;
            }
            if (changed) replacements.Add(structure.Part, XpsPackage.Serialize(structure.Markup));
        }
        return replacements;
    }

    private int[] MapPageOccurrences(IReadOnlyList<XpsPage> next) {
        var positions = new Dictionary<XpsPage, List<int>>();
        for (int i = 0; i < next.Count; i++) {
            if (!positions.TryGetValue(next[i], out var indexes)) positions.Add(next[i], indexes = new List<int>());
            indexes.Add(i);
        }
        var ordinals = new Dictionary<XpsPage, int>();
        var map = new int[_pages.Count];
        for (int i = 0; i < _pages.Count; i++) {
            var page = _pages[i];
            ordinals.TryGetValue(page, out int ordinal); ordinals[page] = ordinal + 1;
            map[i] = positions.TryGetValue(page, out var indexes) ? indexes[Math.Min(ordinal, indexes.Count - 1)] : -1;
        }
        return map;
    }
}

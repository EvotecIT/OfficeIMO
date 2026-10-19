namespace OfficeIMO.Pdf;

internal static partial class PdfStructTreeRootDictionaryBuilder {
    // Only non-root nodes carry Limits; all child nodes are indirect objects.
    internal static string BuildParentTree(IReadOnlyList<ParentTreeEntry> entries, Func<string, int> addObject) {
        Guard.NotNull(entries, nameof(entries));
        Guard.NotNull(addObject, nameof(addObject));
        const int fanout = 64;
        var ordered = entries.OrderBy(entry => entry.StructParentIndex).ToArray();
        for (int i = 1; i < ordered.Length; i++) {
            if (ordered[i - 1].StructParentIndex == ordered[i].StructParentIndex)
                throw new ArgumentException("PDF parent-tree indexes must be unique.", nameof(entries));
        }
        if (ordered.Length <= fanout) return BuildParentTree(ordered);
        var nodes = new List<(int Id, int First, int Last)>();
        for (int start = 0; start < ordered.Length; start += fanout) {
            var leaf = ordered.Skip(start).Take(fanout).ToArray();
            int first = leaf[0].StructParentIndex;
            int last = leaf[leaf.Length - 1].StructParentIndex;
            nodes.Add((addObject(WithLimits(BuildParentTree(leaf), first, last)), first, last));
        }
        while (nodes.Count > fanout) {
            var parents = new List<(int Id, int First, int Last)>();
            for (int start = 0; start < nodes.Count; start += fanout) {
                var group = nodes.Skip(start).Take(fanout).ToArray();
                int first = group[0].First;
                int last = group[group.Length - 1].Last;
                parents.Add((addObject(WithLimits(BuildNumberTreeKids(group.Select(node => node.Id).ToArray()), first, last)), first, last));
            }
            nodes = parents;
        }
        return BuildNumberTreeKids(nodes.Select(node => node.Id).ToArray());
    }

    private static string WithLimits(string dictionary, int first, int last) =>
        dictionary.Substring(0, dictionary.Length - 3) + " /Limits [" +
        first.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " +
        last.ToString(System.Globalization.CultureInfo.InvariantCulture) + "] >>\n";

    private static string BuildNumberTreeKids(IReadOnlyList<int> children) {
        var sb = new StringBuilder("<< /Kids ");
        AppendReferenceArray(sb, children);
        return sb.Append(" >>\n").ToString();
    }

    // Metadata belongs to the structure dictionary, never to a text/content writer.
    internal static string WithStructureMetadata(string dictionary, int? namespaceId = null, string? language = null, string? elementId = null) {
        if (!namespaceId.HasValue && string.IsNullOrWhiteSpace(language) && elementId == null) return dictionary;
        var sb = new StringBuilder(dictionary.Substring(0, dictionary.Length - 3));
        if (namespaceId.HasValue) sb.Append(" /NS ").Append(PdfSyntaxEscaper.IndirectReference(namespaceId.Value));
        if (!string.IsNullOrWhiteSpace(language)) sb.Append(" /Lang ").Append(PdfSyntaxEscaper.TextString(language!));
        if (elementId != null) sb.Append(" /ID ").Append(PdfSyntaxEscaper.TextString(elementId));
        return sb.Append(" >>\n").ToString();
    }
}

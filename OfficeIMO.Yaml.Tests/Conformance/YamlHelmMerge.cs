namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>How the reference merge combines two sequences at the same path.</summary>
public enum YamlHelmSequenceRule {
    /// <summary>The overlay's sequence replaces the base sequence (Helm).</summary>
    Replace,
    /// <summary>The overlay's items follow the base items.</summary>
    Append,
    /// <summary>Mapping items with equal values of a key merge; other overlay items follow the base items.</summary>
    MergeByKey
}

/// <summary>
/// An independent reference for merging loaded values, written from Helm's rules for values files rather than from
/// OfficeIMO.Yaml: maps merge recursively and everything else from the overlay replaces (Helm's <c>mergeMaps</c>),
/// with Helm's convention that an explicit null deletes a key as an option. It works on plain loaded data, so the
/// documents it merges come from another YAML library.
/// </summary>
public static class YamlHelmMerge {
    /// <summary>Merges each overlay document into the base document at the same position.</summary>
    public static IReadOnlyList<YamlConformanceNode> Merge(IReadOnlyList<YamlConformanceNode> baseDocuments, IReadOnlyList<YamlConformanceNode> overlayDocuments, YamlHelmSequenceRule sequences, string? itemKey, bool nullRemovesKey) {
        var result = new List<YamlConformanceNode>(baseDocuments);
        for (int i = 0; i < overlayDocuments.Count; i++) {
            YamlConformanceNode overlay = overlayDocuments[i];
            // An empty overlay file is an empty set of values.
            if (overlay.Kind == YamlConformanceNodeKind.Null) continue;
            if (i < result.Count) result[i] = Value(result[i], overlay, sequences, itemKey, nullRemovesKey);
            else {
                while (result.Count < i) result.Add(YamlConformanceNode.Null);
                result.Add(overlay);
            }
        }

        return result;
    }

    private static YamlConformanceNode Value(YamlConformanceNode current, YamlConformanceNode overlay, YamlHelmSequenceRule sequences, string? itemKey, bool nullRemovesKey) {
        if (current.Kind == YamlConformanceNodeKind.Mapping && overlay.Kind == YamlConformanceNodeKind.Mapping) return Maps(current, overlay, sequences, itemKey, nullRemovesKey);
        if (current.Kind == YamlConformanceNodeKind.Sequence && overlay.Kind == YamlConformanceNodeKind.Sequence) {
            if (sequences == YamlHelmSequenceRule.Append) return YamlConformanceNode.Sequence(current.Items.Concat(overlay.Items).ToList());
            if (sequences == YamlHelmSequenceRule.MergeByKey) return ByKey(current, overlay, itemKey!, sequences, nullRemovesKey);
        }

        return overlay;
    }

    private static YamlConformanceNode Maps(YamlConformanceNode current, YamlConformanceNode overlay, YamlHelmSequenceRule sequences, string? itemKey, bool nullRemovesKey) {
        var entries = current.Entries.ToList();
        foreach (KeyValuePair<YamlConformanceNode, YamlConformanceNode> entry in overlay.Entries) {
            int index = entries.FindIndex(e => SameKey(e.Key, entry.Key));
            if (nullRemovesKey && entry.Value.Kind == YamlConformanceNodeKind.Null) {
                if (index >= 0) entries.RemoveAt(index);
                continue;
            }

            if (index < 0) entries.Add(entry);
            else entries[index] = new KeyValuePair<YamlConformanceNode, YamlConformanceNode>(entries[index].Key, Value(entries[index].Value, entry.Value, sequences, itemKey, nullRemovesKey));
        }

        return YamlConformanceNode.Mapping(entries);
    }

    private static YamlConformanceNode ByKey(YamlConformanceNode current, YamlConformanceNode overlay, string itemKey, YamlHelmSequenceRule sequences, bool nullRemovesKey) {
        var items = current.Items.ToList();
        foreach (YamlConformanceNode item in overlay.Items) {
            string? id = Id(item, itemKey);
            int index = id == null ? -1 : items.FindIndex(i => Id(i, itemKey) == id);
            if (index < 0) items.Add(item);
            else items[index] = Maps(items[index], item, sequences, itemKey, nullRemovesKey);
        }

        return YamlConformanceNode.Sequence(items);
    }

    private static string? Id(YamlConformanceNode item, string key) {
        if (item.Kind != YamlConformanceNodeKind.Mapping) return null;
        foreach (KeyValuePair<YamlConformanceNode, YamlConformanceNode> entry in item.Entries) {
            if (entry.Key.Scalar == key && entry.Value.Kind != YamlConformanceNodeKind.Null && entry.Value.Scalar != null) return entry.Value.Scalar;
        }

        return null;
    }

    // Helm's maps are keyed by strings, so keys compare by their text.
    private static bool SameKey(YamlConformanceNode left, YamlConformanceNode right) =>
        left.Scalar != null && string.Equals(left.Scalar, right.Scalar, StringComparison.Ordinal);
}

namespace OfficeIMO.Yaml;

/// <summary>
/// Compares what two texts mean, not how they are written: the documents, node kinds, decoded scalar values and what
/// they resolve to, tags, anchors and aliases. Comments, styles and layout are ignored.
/// </summary>
internal static class YamlValueEquivalence {
    public static bool SameValues(string before, string after, YamlLoadOptions? options) {
        YamlDocument? a = null, b = null;
        try {
            a = YamlDocument.Parse(before, options);
            b = YamlDocument.Parse(after, options);
            return SameValues(a, b);
        } catch (YamlException) {
            return false;
        } finally {
            a?.Dispose();
            b?.Dispose();
        }
    }

    public static bool SameValues(YamlDocument a, YamlDocument b) {
        IReadOnlyList<YamlDocumentPart> x = a.Documents, y = b.Documents;
        if (x.Count != y.Count) return false;
        for (int i = 0; i < x.Count; i++) {
            YamlNode? p = x[i].RootNode, q = y[i].RootNode;
            if (p is null || q is null) {
                if (!IsEmpty(p) || !IsEmpty(q)) return false;
                continue;
            }
            if (!Same(p, q)) return false;
        }
        return true;
    }

    /// <summary>An absent root and an empty plain scalar are the same empty document.</summary>
    private static bool IsEmpty(YamlNode? node) => node is null || node is YamlScalar { IsNull: true, Tag: null, Anchor: null };

    private static bool Same(YamlNode a, YamlNode b) {
        if (a.GetType() != b.GetType() || a.Tag != b.Tag || a.Anchor != b.Anchor) return false;
        switch (a) {
            case YamlScalar sa: {
                var sb = (YamlScalar)b;
                YamlScalarKind kind = sa.Kind;
                if (kind != sb.Kind) return false;
                // Typed scalars compare by value, so YAML 1.1 'yes' and 'true' are the same boolean.
                switch (kind) {
                    case YamlScalarKind.Null:
                        return true;
                    case YamlScalarKind.Boolean:
                        return sa.TryGetValue(out bool ba) && sb.TryGetValue(out bool bb) && ba == bb;
                    case YamlScalarKind.Integer when sa.TryGetValue(out long la) && sb.TryGetValue(out long lb):
                        return la == lb;
                    case YamlScalarKind.Float when sa.TryGetValue(out double da) && sb.TryGetValue(out double db):
                        return da.Equals(db);
                    default:
                        return string.Equals(sa.Value, sb.Value, StringComparison.Ordinal);
                }
            }
            case YamlAlias aa:
                return aa.Name == ((YamlAlias)b).Name;
            case YamlMapping ma: {
                var mb = (YamlMapping)b;
                if (ma._count != mb._count) return false;
                for (int i = 0; i < ma._count; i++) {
                    if (!Same(ma.Entries[i].Key, mb.Entries[i].Key) || !Same(ma.Entries[i].Value, mb.Entries[i].Value)) return false;
                }
                return true;
            }
            case YamlSequence qa: {
                var qb = (YamlSequence)b;
                if (qa._count != qb._count) return false;
                for (int i = 0; i < qa._count; i++) {
                    if (!Same(qa.Items[i], qb.Items[i])) return false;
                }
                return true;
            }
            default:
                return false;
        }
    }
}

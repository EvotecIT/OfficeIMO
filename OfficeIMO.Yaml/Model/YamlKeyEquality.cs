namespace OfficeIMO.Yaml;

/// <summary>
/// Equality of mapping keys as YAML values (specification section 3.2.1.3): the same resolved type and the same
/// canonical value, so <c>a</c> equals <c>"a"</c>, <c>0x1</c> equals <c>1</c>, and <c>1</c> differs from <c>"1"</c>.
/// Custom-tagged scalars compare by tag and text; collections compare structurally through aliases.
/// </summary>
internal static class YamlKeyEquality {
    private const char Marker = '\u0001';
    private const string NullIdentity = "\u0001null";
    private const string TrueIdentity = "\u0001true";
    private const string FalseIdentity = "\u0001false";

    /// <summary>
    /// Returns a string that is equal for equal scalar keys, or null for collections. String keys return their decoded
    /// text, which the load already cached, so the common case allocates nothing. Unequal keys may share an identity.
    /// </summary>
    public static string? Identity(YamlNode key) {
        if (key.Resolved is not YamlScalar scalar) return null;
        if (!scalar.TryGetKind(out YamlScalarKind kind)) return scalar.Tag + Marker + scalar.Value;
        ReadOnlySpan<char> text = scalar.ValueSpan;
        switch (kind) {
            case YamlScalarKind.Null:
                return NullIdentity;
            case YamlScalarKind.Boolean:
                if (YamlSchemaResolver.TryParseBoolean(text, scalar.NumericSchema, out bool flag)) return flag ? TrueIdentity : FalseIdentity;
                break;
            case YamlScalarKind.Integer:
                if (YamlSchemaResolver.TryParseInt64(text, scalar.NumericSchema, out long integer)) return Marker + integer.ToString(CultureInfo.InvariantCulture);
                break;
            case YamlScalarKind.Float:
                if (YamlSchemaResolver.TryParseDouble(text, scalar.NumericSchema, out double number)) return Marker + number.ToString("R", CultureInfo.InvariantCulture);
                break;
            case YamlScalarKind.String:
                return scalar.Value;
        }
        return Marker + scalar.Value;
    }

    /// <summary>Compares two keys as values. Called only for keys whose identities are equal, or for collection keys.</summary>
    public static bool Equal(YamlNode left, YamlNode right) {
        YamlNode a = left.Resolved, b = right.Resolved;
        if (ReferenceEquals(a, b)) return true;
        switch (a) {
            case YamlScalar sa when b is YamlScalar sb:
                bool ka = sa.TryGetKind(out YamlScalarKind kindA), kb = sb.TryGetKind(out YamlScalarKind kindB);
                if (ka != kb) return false;
                if (!ka) return sa.Tag == sb.Tag && sa.Value == sb.Value;
                // Repeated YAML 1.1 merge keys are merge instructions, not duplicate entries.
                if (kindA != kindB || kindA == YamlScalarKind.Merge) return false;
                // Integers and floats share the "\u0001" prefix; an equal identity with different kinds is not a duplicate.
                return string.Equals(Identity(sa), Identity(sb), StringComparison.Ordinal);
            case YamlSequence qa when b is YamlSequence qb:
                if (qa._count != qb._count || qa.Tag != qb.Tag) return false;
                for (int i = 0; i < qa._count; i++) {
                    if (!Equal(qa.Items[i], qb.Items[i])) return false;
                }
                return true;
            case YamlMapping ma when b is YamlMapping mb:
                if (ma._count != mb._count || ma.Tag != mb.Tag) return false;
                for (int i = 0; i < ma._count; i++) {
                    if (!Equal(ma.Entries[i].Key, mb.Entries[i].Key) || !Equal(ma.Entries[i].Value, mb.Entries[i].Value)) return false;
                }
                return true;
            default:
                return false;
        }
    }
}

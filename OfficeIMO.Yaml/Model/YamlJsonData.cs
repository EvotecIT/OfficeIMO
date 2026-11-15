using System.Runtime.CompilerServices;
using OfficeIMO.Json;

namespace OfficeIMO.Yaml;

/// <summary>
/// Converts the document model to the JSON data model that <see cref="JsonSchema"/> validates. Scalars take the type
/// their document's schema resolves (so <c>yes</c> is a boolean under YAML 1.1 and a string under the core schema),
/// merge keys apply, aliases share their anchored value, and every value keeps its source position and node.
/// </summary>
internal sealed class YamlJsonData {
    private const int MaxDepth = 512;
    private Dictionary<YamlNode, KeyValuePair<JsonData, int>>? _anchored;

    public static JsonData Convert(YamlNode? node) => node is null ? JsonData.CreateNull() : new YamlJsonData().ConvertNode(node, 0, out _);

    /// <summary>Converts a node at <paramref name="depth"/>; <paramref name="height"/> is how many levels its value nests below it.</summary>
    private JsonData ConvertNode(YamlNode node, int depth, out int height) {
        if (depth > MaxDepth) throw TooDeep();
        YamlNode resolved = node.Resolved;
        bool shared = resolved.Anchor is not null;
        // Aliases share one converted value, scalars included, so alias-heavy documents convert in linear time and
        // memory; a shared value keeps the anchored node's position.
        // A cached value still counts its full nesting at the depth it is reused.
        if (shared && _anchored is not null && _anchored.TryGetValue(resolved, out KeyValuePair<JsonData, int> existing)) {
            height = existing.Value;
            if (depth + height > MaxDepth) throw TooDeep();
            return existing.Key;
        }
        height = 0;
        JsonData value = resolved switch {
            YamlScalar scalar => ConvertScalar(scalar, Locate(shared ? resolved : node)),
            YamlSequence sequence => ConvertSequence(sequence, depth, out height),
            YamlMapping mapping => ConvertMapping(mapping, depth, out height),
            _ => JsonData.CreateNull(Locate(node))
        };
        if (shared) (_anchored ??= new Dictionary<YamlNode, KeyValuePair<JsonData, int>>(ReferenceComparer.Instance))[resolved] = new KeyValuePair<JsonData, int>(value, height);
        return value;
    }

    private static InvalidOperationException TooDeep() =>
        new InvalidOperationException("The YAML value nests deeper than " + MaxDepth.ToString(CultureInfo.InvariantCulture) + " levels, or an alias refers to a node that contains it.");

    private static JsonDataLocation Locate(YamlNode node) {
        YamlPosition start = node.Start;
        return new JsonDataLocation(start.Line, start.Column, node);
    }

    private static JsonData ConvertScalar(YamlScalar scalar, JsonDataLocation location) {
        switch (scalar.Kind) {
            case YamlScalarKind.Null:
                if (scalar.IsNull) return JsonData.CreateNull(location);
                break;
            case YamlScalarKind.Boolean:
                if (scalar.TryGetValue(out bool flag)) return JsonData.CreateBoolean(flag, location);
                break;
            case YamlScalarKind.Integer:
                // An explicit !!int tag on text that is not integer syntax, such as 1.5, keeps the text.
                if (!YamlSchemaResolver.IsInteger(scalar.Value.AsSpan(), scalar.Schema)) break;
                if (scalar.TryGetValue(out decimal exact)) return JsonData.CreateNumber(exact, location);
                // Decimal integers beyond decimal range stay exact, as the untyped reader keeps them as BigInteger.
                if (YamlBigIntegerConverter.TryParse(scalar.Value.AsSpan(), out System.Numerics.BigInteger big)) {
                    return JsonData.ParseNumber(big.ToString(CultureInfo.InvariantCulture), location);
                }
                if (scalar.TryGetValue(out double large)) return JsonData.CreateNumber(large, location);
                break;
            case YamlScalarKind.Float:
                if (scalar.TryGetValue(out double number)) return JsonData.CreateNumber(number, location);
                break;
        }
        // Strings, timestamps, custom tags and scalars a tag claims but cannot hold keep their text.
        return JsonData.CreateString(scalar.Value, location);
    }

    private JsonData ConvertSequence(YamlSequence sequence, int depth, out int height) {
        var items = new JsonData[sequence.Count];
        height = 1;
        for (int i = 0; i < items.Length; i++) {
            items[i] = ConvertNode(sequence[i], depth + 1, out int child);
            height = Math.Max(height, child + 1);
        }
        return JsonData.CreateArray(items, Locate(sequence));
    }

    private JsonData ConvertMapping(YamlMapping mapping, int depth, out int height) {
        height = 1;
        var properties = new List<KeyValuePair<string, JsonData>>(mapping.Count);
        var seen = new Dictionary<string, int>(StringComparer.Ordinal);
        bool firstWins = mapping.Document?.DuplicateKeys == YamlDuplicateKeyPolicy.FirstWins;
        List<YamlNode>? merges = null;
        for (int i = 0; i < mapping.Count; i++) {
            if (mapping.IsMergeKeyAt(i)) {
                (merges ??= new List<YamlNode>()).Add(mapping[i].Value);
                continue;
            }
            YamlMappingEntry entry = mapping[i];
            string key = KeyText(entry.Key);
            if (seen.TryGetValue(key, out int index)) {
                if (!firstWins) properties[index] = new KeyValuePair<string, JsonData>(key, ConvertChild(entry.Value, depth, ref height));
                continue;
            }
            seen.Add(key, properties.Count);
            properties.Add(new KeyValuePair<string, JsonData>(key, ConvertChild(entry.Value, depth, ref height)));
        }
        if (merges is not null) {
            // The mapping's own keys win, then merged mappings in order.
            foreach (YamlNode merge in merges) {
                YamlNode source = merge.Resolved;
                if (source is YamlMapping merged) {
                    AddMerged(properties, seen, ConvertChild(merge, depth, ref height));
                } else if (source is YamlSequence list) {
                    foreach (YamlNode item in list) {
                        if (item.Resolved is YamlMapping) AddMerged(properties, seen, ConvertChild(item, depth, ref height));
                    }
                }
            }
        }
        return JsonData.CreateObject(properties, Locate(mapping));
    }

    private JsonData ConvertChild(YamlNode node, int depth, ref int height) {
        JsonData value = ConvertNode(node, depth + 1, out int child);
        height = Math.Max(height, child + 1);
        return value;
    }

    private static void AddMerged(List<KeyValuePair<string, JsonData>> properties, Dictionary<string, int> seen, JsonData merged) {
        foreach (KeyValuePair<string, JsonData> property in merged.Properties) {
            if (seen.ContainsKey(property.Key)) continue;
            seen.Add(property.Key, properties.Count);
            properties.Add(property);
        }
    }

    /// <summary>Gets a key's text: a scalar's value, or the flow text of a collection key, which JSON cannot hold as a name.</summary>
    private static string KeyText(YamlNode key) => key.Resolved is YamlScalar scalar ? scalar.Value : key.Resolved.ToString();

    private sealed class ReferenceComparer : IEqualityComparer<YamlNode> {
        public static readonly ReferenceComparer Instance = new ReferenceComparer();
        public bool Equals(YamlNode? x, YamlNode? y) => ReferenceEquals(x, y);
        public int GetHashCode(YamlNode obj) => RuntimeHelpers.GetHashCode(obj);
    }
}

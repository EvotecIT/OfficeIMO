using System.Globalization;
using Microsoft.Extensions.Configuration;

namespace OfficeIMO.Yaml.Configuration;

/// <summary>
/// Flattens one YAML document into configuration keys: mapping keys and sequence indexes joined with <c>:</c>, aliases and
/// merge keys resolved, nulls as empty strings, and empty collections as <see langword="null"/>, as the JSON provider does.
/// </summary>
internal sealed class YamlConfigurationFileParser {
    private readonly Dictionary<string, string?> _data = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

    private YamlConfigurationFileParser() { }

    public static IDictionary<string, string?> Parse(Stream input) {
        YamlDocument document;
        try {
            document = YamlDocument.Load(input);
        } catch (YamlException exception) {
            throw new FormatException(string.Format(CultureInfo.InvariantCulture,
                "Could not parse the YAML configuration at line {0}, column {1}: {2}", exception.Line, exception.Column, exception.Message), exception);
        }

        using (document) {
            var parser = new YamlConfigurationFileParser();
            parser.VisitDocument(document);
            return parser._data;
        }
    }

    private void VisitDocument(YamlDocument document) {
        IReadOnlyList<YamlDocumentPart> parts = document.Documents;
        if (parts.Count > 1) {
            int line = parts[1].Root.Start.Line;
            throw new FormatException(string.Format(CultureInfo.InvariantCulture,
                "A YAML configuration file holds one document, but this stream has {0}; the second starts at line {1}. Split the documents into separate files.",
                parts.Count, line));
        }
        if (parts.Count == 0) return;

        YamlNode root = Resolve(parts[0].Root);
        if (root is YamlMapping mapping) {
            VisitMapping(mapping, path: null);
        } else if (!root.IsNull) {
            throw new FormatException(string.Format(CultureInfo.InvariantCulture,
                "The root of a YAML configuration file must be a mapping, but line {0} starts {1}.", root.Start.Line, Describe(root)));
        }
    }

    private void VisitNode(YamlNode node, string path) {
        node = Resolve(node);
        switch (node) {
            case YamlMapping mapping:
                VisitMapping(mapping, path);
                break;
            case YamlSequence sequence:
                if (sequence.Count == 0) {
                    Set(path, null, sequence);
                    break;
                }
                for (int i = 0; i < sequence.Count; i++) {
                    VisitNode(sequence[i], Combine(path, i.ToString(CultureInfo.InvariantCulture)));
                }
                break;
            default:
                Set(path, node.IsNull ? string.Empty : node.AsString(), node);
                break;
        }
    }

    // The root mapping has no path; an empty key ("") is a real segment, so a null path marks the root.
    private void VisitMapping(YamlMapping mapping, string? path) {
        var resolved = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        VisitEntries(mapping, path, resolved, merged: false);
        if (resolved.Count == 0 && path is not null) Set(path, null, mapping);
    }

    // Different YAML paths can flatten to one configuration key ("a:b": 1 beside a: {b: []}), so every write checks.
    private void Set(string path, string? value, YamlNode node) {
        if (_data.ContainsKey(path)) throw Duplicate(path, node);
        _data[path] = value;
    }

    // A mapping's own keys win over the keys merged into it, and earlier merged mappings win over later ones, so every key
    // is taken from the first mapping that sets it. Duplicates within one mapping are errors wherever the mapping is used.
    private void VisitEntries(YamlMapping mapping, string? path, HashSet<string> resolved, bool merged) {
        HashSet<string>? own = merged ? new HashSet<string>(StringComparer.OrdinalIgnoreCase) : null;
        bool hasMerges = false;
        foreach (YamlMappingEntry entry in mapping) {
            if (Resolve(entry.Key) is not YamlScalar key) {
                throw new FormatException(string.Format(CultureInfo.InvariantCulture,
                    "Configuration keys must be scalars, but the key at line {0} is {1}.", entry.Key.Start.Line, Describe(Resolve(entry.Key))));
            }
            if (key.Kind == YamlScalarKind.Merge) {
                hasMerges = true;
                continue;
            }

            string name = key.AsString();
            if (own is not null && !own.Add(name)) throw Duplicate(Combine(path, name), key);
            if (!resolved.Add(name)) {
                if (merged) continue;
                throw Duplicate(Combine(path, name), key);
            }
            VisitNode(entry.Value, Combine(path, name));
        }

        if (!hasMerges) return;
        foreach (YamlMappingEntry entry in mapping) {
            if (Resolve(entry.Key) is not YamlScalar { Kind: YamlScalarKind.Merge }) continue;
            YamlNode source = Resolve(entry.Value);
            if (source is YamlMapping single) {
                VisitEntries(single, path, resolved, merged: true);
            } else if (source is YamlSequence list) {
                foreach (YamlNode item in list) {
                    if (Resolve(item) is not YamlMapping itemMapping) throw InvalidMerge(item);
                    VisitEntries(itemMapping, path, resolved, merged: true);
                }
            } else if (!source.IsNull) {
                throw InvalidMerge(source);
            }
        }
    }

    private static YamlNode Resolve(YamlNode node) {
        while (node is YamlAlias alias) node = alias.Target;
        return node;
    }

    private static string Combine(string? path, string key) => path is null ? key : path + ConfigurationPath.KeyDelimiter + key;

    private static string Describe(YamlNode node) => node switch {
        YamlMapping => "a mapping",
        YamlSequence => "a sequence",
        _ => "a scalar"
    };

    private static FormatException Duplicate(string key, YamlNode node) =>
        new FormatException(string.Format(CultureInfo.InvariantCulture, "A duplicate key '{0}' was found at line {1}.", key, node.Start.Line));

    private static FormatException InvalidMerge(YamlNode node) =>
        new FormatException(string.Format(CultureInfo.InvariantCulture,
            "A merge key (<<) takes a mapping or a sequence of mappings, but line {0} gives {1}.", node.Start.Line, Describe(Resolve(node))));
}

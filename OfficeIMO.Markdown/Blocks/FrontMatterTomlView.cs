using OfficeIMO.Toml;
using OfficeIMO.Yaml;

namespace OfficeIMO.Markdown;

/// <summary>
/// Builds the YAML view through which TOML front matter is read, so both front matter languages share one value model.
/// YAML's data model holds every TOML value: tables become mappings, arrays and arrays of tables become sequences, and
/// dates and times keep their TOML spelling as plain scalars.
/// </summary>
internal static class FrontMatterTomlView {
    public static YamlMapping ToMapping(TomlTable table) {
        var mapping = new YamlMapping();
        foreach (KeyValuePair<string, TomlNode> entry in table) mapping.Add(entry.Key, ToNode(entry.Value));
        return mapping;
    }

    private static YamlNode ToNode(TomlNode node) {
        switch (node) {
            case TomlTable table:
                return ToMapping(table);
            case TomlTableArray tables: {
                var sequence = new YamlSequence();
                foreach (TomlTable item in tables) sequence.Add(ToMapping(item));
                return sequence;
            }
            case TomlArray array: {
                var sequence = new YamlSequence();
                foreach (TomlNode item in array) sequence.Add(ToNode(item));
                return sequence;
            }
            default:
                TomlValue value = node.AsValue();
                return value.Kind switch {
                    // Quoted, so a TOML string such as "5" or "true" stays a string under the YAML schema.
                    TomlNodeKind.String => new YamlScalar(value.AsString(), YamlStyle.DoubleQuoted),
                    TomlNodeKind.Integer => new YamlScalar(value.GetValue<long>()),
                    TomlNodeKind.Float => new YamlScalar(value.GetValue<double>()),
                    TomlNodeKind.Boolean => new YamlScalar(value.GetValue<bool>()),
                    _ => new YamlScalar(value.AsString(), YamlStyle.Plain)
                };
        }
    }
}

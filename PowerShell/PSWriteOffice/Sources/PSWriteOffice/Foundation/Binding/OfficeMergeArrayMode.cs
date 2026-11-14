using System.Collections;

namespace PSWriteOffice.Foundation;

/// <summary>
/// Binds the array mode of a YAML or TOML merge from the forms a script can write: 'Replace', 'Append', or
/// @{ MergeByKey = 'name' }. The native mode types have only factory members, which scripts cannot reach by type name.
/// </summary>
internal static class OfficeMergeArrayMode {
    internal static object Convert(Type type, object? input) {
        if (input != null && type.IsInstanceOfType(input)) return input;
        bool toml = type == typeof(OfficeIMO.Toml.TomlArrayMerge);
        if (input is string name) {
            if (name.Equals("Replace", StringComparison.OrdinalIgnoreCase)) return toml ? OfficeIMO.Toml.TomlArrayMerge.Replace : OfficeIMO.Yaml.YamlSequenceMerge.Replace;
            if (name.Equals("Append", StringComparison.OrdinalIgnoreCase)) return toml ? OfficeIMO.Toml.TomlArrayMerge.Append : OfficeIMO.Yaml.YamlSequenceMerge.Append;
        } else if (input is IDictionary table && table.Count == 1) {
            foreach (DictionaryEntry entry in table) {
                if (entry.Key is string key && key.Equals("MergeByKey", StringComparison.OrdinalIgnoreCase) && OfficeArgumentValue.Unwrap(entry.Value) is string keyName) {
                    return toml ? OfficeIMO.Toml.TomlArrayMerge.MergeByKey(keyName) : OfficeIMO.Yaml.YamlSequenceMerge.MergeByKey(keyName);
                }
            }
        }
        throw new ArgumentException("Use 'Replace', 'Append' or @{ MergeByKey = 'name' }.");
    }
}

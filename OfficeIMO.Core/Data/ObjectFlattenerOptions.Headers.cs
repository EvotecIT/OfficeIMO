using System;

namespace OfficeIMO.Data {
    public partial class ObjectFlattenerOptions {
        /// <summary>Resolves display prefixes without changing paths used to select or format values.</summary>
        internal string GetHeaderPath(string path) {
            string? matchingCollection = null;
            CollectionColumnMapping? matchingMapping = null;
            foreach (var mapping in CollectionMapColumns) {
                if (mapping.Value.HeaderPrefix == null ||
                    !path.StartsWith(mapping.Key + ".", StringComparison.OrdinalIgnoreCase)) continue;
                if (matchingCollection == null || mapping.Key.Length > matchingCollection.Length) {
                    matchingCollection = mapping.Key;
                    matchingMapping = mapping.Value;
                }
            }
            if (matchingCollection != null) {
                string suffix = path.Substring(matchingCollection.Length + 1);
                string prefix = matchingMapping!.HeaderPrefix!;
                path = prefix.Length == 0 ? suffix : prefix.TrimEnd('.') + "." + suffix;
            }
            foreach (string prefix in HeaderPrefixTrimPaths) {
                if (!string.IsNullOrEmpty(prefix) && path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
                    path = path.Substring(prefix.Length);
                }
            }
            return path;
        }
    }
}

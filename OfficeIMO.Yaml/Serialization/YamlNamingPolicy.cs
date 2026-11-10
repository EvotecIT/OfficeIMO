namespace OfficeIMO.Yaml;

/// <summary>
/// Converts .NET property names to mapping keys. The built-in policies cover the conventions YAML files use:
/// <see cref="CamelCase"/> (Kubernetes, OpenAPI), <see cref="SnakeCaseLower"/> (Compose, Ansible) and
/// <see cref="KebabCaseLower"/> (GitHub Actions). Derive from it for another convention.
/// </summary>
public abstract class YamlNamingPolicy {
    /// <summary>Creates a policy.</summary>
    protected YamlNamingPolicy() {
    }

    /// <summary>Gets the policy that writes <c>imagePullPolicy</c> for <c>ImagePullPolicy</c>.</summary>
    public static YamlNamingPolicy CamelCase { get; } = new CamelCasePolicy();

    /// <summary>Gets the policy that writes <c>image_pull_policy</c> for <c>ImagePullPolicy</c>.</summary>
    public static YamlNamingPolicy SnakeCaseLower { get; } = new SeparatedPolicy('_');

    /// <summary>Gets the policy that writes <c>image-pull-policy</c> for <c>ImagePullPolicy</c>.</summary>
    public static YamlNamingPolicy KebabCaseLower { get; } = new SeparatedPolicy('-');

    /// <summary>Returns the mapping key for a property name.</summary>
    /// <param name="name">The .NET property name.</param>
    public abstract string ConvertName(string name);

    private sealed class CamelCasePolicy : YamlNamingPolicy {
        public override string ConvertName(string name) {
            if (string.IsNullOrEmpty(name) || !char.IsUpper(name[0])) return name;
            char[] chars = name.ToCharArray();
            // Lower the leading run of capitals, keeping the last one of a run that a lower-case letter follows
            // (URLValue -> urlValue), as System.Text.Json does.
            for (int i = 0; i < chars.Length; i++) {
                if (i == 1 && !char.IsUpper(chars[i])) break;
                bool hasNext = i + 1 < chars.Length;
                if (i > 0 && hasNext && !char.IsUpper(chars[i + 1])) {
                    if (char.IsSeparator(chars[i + 1])) chars[i] = char.ToLowerInvariant(chars[i]);
                    break;
                }
                chars[i] = char.ToLowerInvariant(chars[i]);
            }
            return new string(chars);
        }
    }

    private sealed class SeparatedPolicy : YamlNamingPolicy {
        private readonly char _separator;

        public SeparatedPolicy(char separator) {
            _separator = separator;
        }

        public override string ConvertName(string name) {
            if (string.IsNullOrEmpty(name)) return name;
            var text = new StringBuilder(name.Length + 8);
            // A word starts at an upper-case letter after a lower-case letter or digit, at the last capital of a run
            // followed by a lower-case letter (HTTPServer -> http_server), and after a separator in the source.
            for (int i = 0; i < name.Length; i++) {
                char c = name[i];
                if (c == '_' || c == '-' || c == ' ') {
                    if (text.Length > 0 && text[text.Length - 1] != _separator) text.Append(_separator);
                    continue;
                }
                if (char.IsUpper(c)) {
                    bool boundary = i > 0 && (char.IsLower(name[i - 1]) || char.IsDigit(name[i - 1])
                        || (char.IsUpper(name[i - 1]) && i + 1 < name.Length && char.IsLower(name[i + 1])));
                    if (boundary && text.Length > 0 && text[text.Length - 1] != _separator) text.Append(_separator);
                    text.Append(char.ToLowerInvariant(c));
                } else {
                    text.Append(c);
                }
            }
            return text.ToString();
        }
    }
}

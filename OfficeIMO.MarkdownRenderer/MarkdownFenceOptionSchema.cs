using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Declares a host/plugin fence metadata schema for one or more fenced languages.
/// </summary>
public sealed class MarkdownFenceOptionSchema {
    private static readonly HashSet<string> CoreMetadataAliases = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "title",
        "caption",
        "id",
        "class"
    };

    private readonly IReadOnlyList<string> _languages;
    private readonly IReadOnlyList<MarkdownFenceOptionDefinition> _definitions;

    /// <summary>
    /// Creates a new fence option schema.
    /// </summary>
    public MarkdownFenceOptionSchema(
        string id,
        string name,
        IEnumerable<string> languages,
        IEnumerable<MarkdownFenceOptionDefinition> definitions) {
        if (string.IsNullOrWhiteSpace(id)) {
            throw new ArgumentException("Schema id is required.", nameof(id));
        }

        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Schema name is required.", nameof(name));
        }

        if (languages == null) {
            throw new ArgumentNullException(nameof(languages));
        }

        if (definitions == null) {
            throw new ArgumentNullException(nameof(definitions));
        }

        Id = id.Trim();
        Name = name.Trim();
        _languages = NormalizeValues(languages);
        _definitions = NormalizeDefinitions(definitions);

        if (_languages.Count == 0) {
            throw new ArgumentException("At least one language is required.", nameof(languages));
        }

        if (_definitions.Count == 0) {
            throw new ArgumentException("At least one option definition is required.", nameof(definitions));
        }
    }

    /// <summary>Stable schema identifier used for registration and diagnostics.</summary>
    public string Id { get; }

    /// <summary>Friendly schema name.</summary>
    public string Name { get; }

    /// <summary>Fence languages handled by this schema.</summary>
    public IReadOnlyList<string> Languages => _languages;

    /// <summary>All supported option definitions.</summary>
    public IReadOnlyList<MarkdownFenceOptionDefinition> Definitions => _definitions;

    /// <summary>
    /// Returns <see langword="true"/> when the schema handles the supplied language.
    /// </summary>
    public bool HandlesLanguage(string language) {
        if (string.IsNullOrWhiteSpace(language)) {
            return false;
        }

        for (int i = 0; i < _languages.Count; i++) {
            if (string.Equals(_languages[i], language.Trim(), StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Parses the supplied fence metadata according to this schema.
    /// </summary>
    public MarkdownFenceOptionSet Parse(MarkdownCodeFenceInfo? fenceInfo) {
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        var errors = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var unknownOptions = new List<string>();

        var attributes = fenceInfo?.Attributes;
        if (attributes != null) {
            foreach (var attribute in attributes) {
                if (string.IsNullOrWhiteSpace(attribute.Key)) {
                    continue;
                }

                if (TryResolveDefinition(attribute.Key, out var definition)) {
                    var error = definition.Validate(attribute.Value);
                    if (!string.IsNullOrWhiteSpace(error)) {
                        errors[definition.Name] = error!;
                        continue;
                    }

                    values[definition.Name] = attribute.Value;
                    continue;
                }

                if (!CoreMetadataAliases.Contains(attribute.Key)) {
                    AddUnknownOption(unknownOptions, attribute.Key);
                }
            }
        }

        return new MarkdownFenceOptionSet(this, fenceInfo, values, errors, unknownOptions.AsReadOnly());
    }

    internal bool TryResolveDefinition(string candidate, out MarkdownFenceOptionDefinition definition) {
        if (string.IsNullOrWhiteSpace(candidate)) {
            definition = null!;
            return false;
        }

        for (int i = 0; i < _definitions.Count; i++) {
            if (_definitions[i].Matches(candidate)) {
                definition = _definitions[i];
                return true;
            }
        }

        definition = null!;
        return false;
    }

    private static IReadOnlyList<string> NormalizeValues(IEnumerable<string> values) {
        var normalized = new List<string>();
        foreach (var value in values) {
            if (string.IsNullOrWhiteSpace(value)) {
                continue;
            }

            bool exists = false;
            var candidate = value.Trim();
            for (int i = 0; i < normalized.Count; i++) {
                if (string.Equals(normalized[i], candidate, StringComparison.OrdinalIgnoreCase)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                normalized.Add(candidate);
            }
        }

        return normalized.AsReadOnly();
    }

    private static IReadOnlyList<MarkdownFenceOptionDefinition> NormalizeDefinitions(IEnumerable<MarkdownFenceOptionDefinition> definitions) {
        var normalized = new List<MarkdownFenceOptionDefinition>();
        foreach (var definition in definitions) {
            if (definition == null) {
                continue;
            }

            bool exists = false;
            for (int i = 0; i < normalized.Count; i++) {
                if (string.Equals(normalized[i].Name, definition.Name, StringComparison.OrdinalIgnoreCase)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                normalized.Add(definition);
            }
        }

        return normalized.AsReadOnly();
    }

    private static void AddUnknownOption(ICollection<string> unknownOptions, string optionName) {
        if (string.IsNullOrWhiteSpace(optionName)) {
            return;
        }

        foreach (var existing in unknownOptions) {
            if (string.Equals(existing, optionName.Trim(), StringComparison.OrdinalIgnoreCase)) {
                return;
            }
        }

        unknownOptions.Add(optionName.Trim());
    }
}

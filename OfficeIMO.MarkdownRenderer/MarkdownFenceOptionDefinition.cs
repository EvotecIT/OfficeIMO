namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Declares one supported fenced metadata option for a host/plugin schema.
/// </summary>
public sealed class MarkdownFenceOptionDefinition {
    private readonly IReadOnlyList<string> _aliases;

    /// <summary>
    /// Creates a new fence option definition.
    /// </summary>
    public MarkdownFenceOptionDefinition(
        string name,
        MarkdownFenceOptionValueKind valueKind,
        IEnumerable<string>? aliases = null,
        Func<string?, string?>? validator = null) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Option name is required.", nameof(name));
        }

        Name = name.Trim();
        ValueKind = valueKind;
        Validator = validator;

        var normalizedAliases = new List<string>();
        AddAlias(normalizedAliases, Name);
        if (aliases != null) {
            foreach (var alias in aliases) {
                AddAlias(normalizedAliases, alias);
            }
        }

        _aliases = normalizedAliases.AsReadOnly();
    }

    /// <summary>Canonical option name used in parsed results.</summary>
    public string Name { get; }

    /// <summary>Expected metadata value kind.</summary>
    public MarkdownFenceOptionValueKind ValueKind { get; }

    /// <summary>All supported names, including the canonical name.</summary>
    public IReadOnlyList<string> Aliases => _aliases;

    internal Func<string?, string?>? Validator { get; }

    /// <summary>
    /// Creates a string-valued option definition.
    /// </summary>
    public static MarkdownFenceOptionDefinition String(
        string name,
        IEnumerable<string>? aliases = null,
        Func<string?, string?>? validator = null) =>
        new MarkdownFenceOptionDefinition(name, MarkdownFenceOptionValueKind.String, aliases, validator);

    /// <summary>
    /// Creates a boolean-valued option definition.
    /// </summary>
    public static MarkdownFenceOptionDefinition Boolean(
        string name,
        IEnumerable<string>? aliases = null,
        Func<string?, string?>? validator = null) =>
        new MarkdownFenceOptionDefinition(name, MarkdownFenceOptionValueKind.Boolean, aliases, validator);

    /// <summary>
    /// Creates an integer-valued option definition.
    /// </summary>
    public static MarkdownFenceOptionDefinition Int32(
        string name,
        IEnumerable<string>? aliases = null,
        Func<string?, string?>? validator = null) =>
        new MarkdownFenceOptionDefinition(name, MarkdownFenceOptionValueKind.Int32, aliases, validator);

    internal bool Matches(string candidate) {
        if (string.IsNullOrWhiteSpace(candidate)) {
            return false;
        }

        for (int i = 0; i < _aliases.Count; i++) {
            if (string.Equals(_aliases[i], candidate.Trim(), StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    internal string? Validate(string? rawValue) {
        switch (ValueKind) {
            case MarkdownFenceOptionValueKind.Boolean:
                if (!TryParseBoolean(rawValue, out _)) {
                    return "Expected a boolean value.";
                }
                break;
            case MarkdownFenceOptionValueKind.Int32:
                if (!int.TryParse(rawValue, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out _)) {
                    return "Expected a 32-bit integer value.";
                }
                break;
        }

        return Validator?.Invoke(rawValue);
    }

    private static void AddAlias(ICollection<string> aliases, string? alias) {
        if (string.IsNullOrWhiteSpace(alias)) {
            return;
        }

        var normalized = alias!.Trim();
        foreach (var existing in aliases) {
            if (string.Equals(existing, normalized, StringComparison.OrdinalIgnoreCase)) {
                return;
            }
        }

        aliases.Add(normalized);
    }

    private static bool TryParseBoolean(string? rawValue, out bool value) {
        value = false;
        if (string.IsNullOrWhiteSpace(rawValue)) {
            return false;
        }

        switch (rawValue!.Trim().ToLowerInvariant()) {
            case "true":
            case "1":
            case "yes":
            case "on":
                value = true;
                return true;
            case "false":
            case "0":
            case "no":
            case "off":
                value = false;
                return true;
            default:
                return false;
        }
    }
}

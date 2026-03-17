using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Parsed fenced metadata values for a specific option schema.
/// </summary>
public sealed class MarkdownFenceOptionSet {
    private readonly IReadOnlyDictionary<string, string?> _values;
    private readonly IReadOnlyDictionary<string, string> _errors;
    private readonly IReadOnlyList<string> _unknownOptions;

    internal MarkdownFenceOptionSet(
        MarkdownFenceOptionSchema schema,
        MarkdownCodeFenceInfo? fenceInfo,
        IReadOnlyDictionary<string, string?> values,
        IReadOnlyDictionary<string, string> errors,
        IReadOnlyList<string> unknownOptions) {
        Schema = schema ?? throw new ArgumentNullException(nameof(schema));
        FenceInfo = fenceInfo;
        _values = values ?? new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        _errors = errors ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        _unknownOptions = unknownOptions ?? Array.Empty<string>();
    }

    /// <summary>Schema used for parsing.</summary>
    public MarkdownFenceOptionSchema Schema { get; }

    /// <summary>Original fenced-code info descriptor.</summary>
    public MarkdownCodeFenceInfo? FenceInfo { get; }

    /// <summary>Canonical parsed option values keyed by option name.</summary>
    public IReadOnlyDictionary<string, string?> Values => _values;

    /// <summary>Validation errors keyed by canonical option name.</summary>
    public IReadOnlyDictionary<string, string> Errors => _errors;

    /// <summary>Unknown non-core attributes encountered during schema parsing.</summary>
    public IReadOnlyList<string> UnknownOptions => _unknownOptions;

    /// <summary>Returns <see langword="true"/> when no schema validation errors were found.</summary>
    public bool IsValid => _errors.Count == 0;

    /// <summary>
    /// Returns <see langword="true"/> when the parsed set contains the specified option.
    /// </summary>
    public bool HasOption(string name) {
        return TryResolveValue(name, out _);
    }

    /// <summary>
    /// Attempts to read a string-valued option.
    /// </summary>
    public bool TryGetString(string name, out string? value) {
        return TryResolveValue(name, out value);
    }

    /// <summary>
    /// Attempts to read a boolean-valued option.
    /// </summary>
    public bool TryGetBoolean(string name, out bool value) {
        value = false;
        return TryResolveValue(name, out var rawValue)
            && TryParseBoolean(rawValue, out value);
    }

    /// <summary>
    /// Attempts to read a 32-bit integer-valued option.
    /// </summary>
    public bool TryGetInt32(string name, out int value) {
        value = 0;
        return TryResolveValue(name, out var rawValue)
            && int.TryParse(rawValue, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out value);
    }

    private bool TryResolveValue(string name, out string? value) {
        value = null;
        if (string.IsNullOrWhiteSpace(name) || !Schema.TryResolveDefinition(name, out var definition)) {
            return false;
        }

        return _values.TryGetValue(definition.Name, out value);
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

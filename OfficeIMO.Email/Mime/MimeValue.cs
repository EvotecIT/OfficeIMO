namespace OfficeIMO.Email;

internal sealed class MimeValue {
    private readonly Dictionary<string, string> _parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    internal MimeValue(string value) {
        Value = value;
    }

    internal string Value { get; }

    internal IDictionary<string, string> Parameters => _parameters;

    internal string? GetParameter(string name) {
        string value;
        return _parameters.TryGetValue(name, out value!) ? value : null;
    }
}

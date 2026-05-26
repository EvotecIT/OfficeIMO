namespace OfficeIMO.Pdf;

/// <summary>
/// Simple viewer preference entries discovered from the document catalog.
/// </summary>
public sealed class PdfViewerPreferences {
    internal PdfViewerPreferences(IReadOnlyDictionary<string, string> values) {
        Values = values;
    }

    /// <summary>Simple viewer preference values keyed by PDF preference name.</summary>
    public IReadOnlyDictionary<string, string> Values { get; }

    /// <summary>Number of simple viewer preference entries read from the catalog.</summary>
    public int Count => Values.Count;

    /// <summary>Gets a simple viewer preference value by PDF preference name.</summary>
    public string? GetValue(string name) {
        if (string.IsNullOrEmpty(name)) {
            return null;
        }

        return Values.TryGetValue(name, out var value) ? value : null;
    }

    /// <summary>Gets a boolean viewer preference value by PDF preference name.</summary>
    public bool? GetBoolean(string name) {
        string? value = GetValue(name);
        if (string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        if (string.Equals(value, "false", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        return null;
    }
}

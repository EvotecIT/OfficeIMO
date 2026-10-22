namespace OfficeIMO.Word;

/// <summary>A metric card's value, label, qualifier and optional semantic color.</summary>
public sealed class WordPaperMetric {
    /// <summary>Creates a metric card.</summary>
    /// <param name="label">Small label describing the value.</param>
    /// <param name="value">Formatted value displayed prominently.</param>
    /// <param name="qualifier">Optional context below the label.</param>
    /// <param name="color">Optional semantic RGB value without a leading hash.</param>
    public WordPaperMetric(string label, string value, string? qualifier = null, string? color = null) {
        Label = label;
        Value = value;
        Qualifier = qualifier;
        Color = color;
    }
    /// <summary>Gets the metric label.</summary>
    public string Label { get; }
    /// <summary>Gets the display value.</summary>
    public string Value { get; }
    /// <summary>Gets the explanatory qualifier.</summary>
    public string? Qualifier { get; }
    /// <summary>Gets the optional value color.</summary>
    public string? Color { get; }
}

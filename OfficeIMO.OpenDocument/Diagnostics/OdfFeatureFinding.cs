namespace OfficeIMO.OpenDocument;

/// <summary>One detected OpenDocument feature and its support level.</summary>
public sealed class OdfFeatureFinding {
    /// <summary>Creates a feature finding.</summary>
    public OdfFeatureFinding(string name, OdfFeatureSupport support, string? partPath = null, int count = 1) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Feature name cannot be empty.", nameof(name));
        if (count < 1) throw new ArgumentOutOfRangeException(nameof(count));
        Name = name;
        Support = support;
        PartPath = partPath;
        Count = count;
    }

    /// <summary>Stable feature name.</summary>
    public string Name { get; }
    /// <summary>Current support level.</summary>
    public OdfFeatureSupport Support { get; }
    /// <summary>Package part containing the feature.</summary>
    public string? PartPath { get; }
    /// <summary>Number of detected occurrences.</summary>
    public int Count { get; }
}

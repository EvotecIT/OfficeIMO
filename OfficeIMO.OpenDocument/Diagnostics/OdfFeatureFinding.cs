namespace OfficeIMO.OpenDocument;

/// <summary>One detected OpenDocument feature and its support level.</summary>
public sealed class OdfFeatureFinding {
    /// <summary>Creates a feature finding.</summary>
    public OdfFeatureFinding(string name, OdfFeatureSupport support, string? partPath = null, int count = 1) {
        Name = name ?? throw new ArgumentNullException(nameof(name));
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

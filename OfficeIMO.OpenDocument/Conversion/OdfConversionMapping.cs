namespace OfficeIMO.OpenDocument;

/// <summary>One feature-level result produced by an explicit OpenDocument conversion adapter.</summary>
public sealed class OdfConversionMapping {
    /// <summary>Creates a feature mapping result.</summary>
    public OdfConversionMapping(string feature, OdfConversionMappingStatus status, int count = 1, string? message = null) {
        if (string.IsNullOrWhiteSpace(feature)) throw new ArgumentException("Feature name cannot be empty.", nameof(feature));
        if (count < 1) throw new ArgumentOutOfRangeException(nameof(count));
        Feature = feature;
        Status = status;
        Count = count;
        Message = message;
    }

    /// <summary>Stable feature identifier, such as <c>paragraphs</c> or <c>formulas</c>.</summary>
    public string Feature { get; }
    /// <summary>How the adapter handled the feature.</summary>
    public OdfConversionMappingStatus Status { get; }
    /// <summary>Number of source items represented by this result.</summary>
    public int Count { get; }
    /// <summary>Optional mapping detail or limitation.</summary>
    public string? Message { get; }
}

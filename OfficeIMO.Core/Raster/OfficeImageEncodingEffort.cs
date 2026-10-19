namespace OfficeIMO.Drawing;

/// <summary>Trades encoding work for lossless compression density without changing decoded samples.</summary>
public enum OfficeImageEncodingEffort {
    /// <summary>Uses a single inexpensive compression plan.</summary>
    Fast,
    /// <summary>Uses bounded heuristics and one compression pass; the default for document image exports.</summary>
    Balanced,
    /// <summary>Searches competing filters and transform plans for smaller output.</summary>
    Smallest
}

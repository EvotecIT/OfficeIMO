namespace OfficeIMO.Drawing;

/// <summary>
/// Stroke ending style for open vector paths.
/// </summary>
public enum OfficeStrokeLineCap {
    /// <summary>Flat line ending at the endpoint.</summary>
    Butt,

    /// <summary>Rounded line ending centered on the endpoint.</summary>
    Round,

    /// <summary>Square line ending extending beyond the endpoint.</summary>
    Square
}

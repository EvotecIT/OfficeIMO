namespace OfficeIMO.Drawing;

/// <summary>
/// Represents one measured line produced by the shared dependency-free text layout helper.
/// </summary>
public readonly struct OfficeTextLine {
    /// <summary>
    /// Creates a measured text line.
    /// </summary>
    /// <param name="text">Line text.</param>
    /// <param name="width">Measured line width in the same coordinate space as the caller's font size.</param>
    /// <param name="offsetX">Additional x offset applied to this line inside the text frame.</param>
    public OfficeTextLine(string text, double width, double offsetX = 0D) {
        Text = text ?? string.Empty;
        Width = width;
        OffsetX = offsetX;
    }

    /// <summary>Line text.</summary>
    public string Text { get; }

    /// <summary>Measured line width.</summary>
    public double Width { get; }

    /// <summary>Additional x offset applied to this line inside the text frame.</summary>
    public double OffsetX { get; }
}

using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Represents a measured multi-line text block produced by the shared text layout engine.
/// </summary>
public sealed class OfficeTextBlockLayout {
    /// <summary>
    /// Creates a measured text block.
    /// </summary>
    /// <param name="lines">Measured text lines.</param>
    /// <param name="fontSize">Resolved font size used for measurement.</param>
    /// <param name="lineHeight">Resolved line height.</param>
    /// <param name="width">Measured block width.</param>
    /// <param name="height">Measured block height.</param>
    /// <param name="clipped">Whether the block was clipped or ellipsized to fit a requested bound.</param>
    public OfficeTextBlockLayout(IReadOnlyList<OfficeTextLine> lines, double fontSize, double lineHeight, double width, double height, bool clipped = false) {
        Lines = lines ?? new[] { new OfficeTextLine(string.Empty, 0D) };
        FontSize = fontSize;
        LineHeight = lineHeight;
        Width = width;
        Height = height;
        Clipped = clipped;
    }

    /// <summary>Measured text lines.</summary>
    public IReadOnlyList<OfficeTextLine> Lines { get; }

    /// <summary>Resolved font size used for measurement.</summary>
    public double FontSize { get; }

    /// <summary>Resolved line height.</summary>
    public double LineHeight { get; }

    /// <summary>Measured block width.</summary>
    public double Width { get; }

    /// <summary>Measured block height.</summary>
    public double Height { get; }

    /// <summary>Whether the block was clipped or ellipsized to fit a requested bound.</summary>
    public bool Clipped { get; }
}

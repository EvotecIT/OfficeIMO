using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes one measured line of rich text.
/// </summary>
public sealed class OfficeRichTextLine {
    /// <summary>
    /// Creates a measured rich text line.
    /// </summary>
    /// <param name="segments">Measured segments on the line.</param>
    /// <param name="lineHeight">Optional resolved height for this line. A value of zero lets block renderers fall back to the containing layout line height.</param>
    /// <param name="offsetX">Additional x offset applied to this line inside the text frame.</param>
    public OfficeRichTextLine(IReadOnlyList<OfficeRichTextSegment> segments, double lineHeight = 0D, double offsetX = 0D) {
        Segments = segments ?? Array.Empty<OfficeRichTextSegment>();
        double width = 0D;
        double fontSize = 0D;
        for (int i = 0; i < Segments.Count; i++) {
            width += Segments[i].Width;
            fontSize = Math.Max(fontSize, Segments[i].FontSize);
        }

        Width = width;
        FontSize = fontSize;
        LineHeight = lineHeight > 0D && !double.IsNaN(lineHeight) && !double.IsInfinity(lineHeight) ? lineHeight : 0D;
        OffsetX = offsetX > 0D && !double.IsNaN(offsetX) && !double.IsInfinity(offsetX) ? offsetX : 0D;
    }

    /// <summary>
    /// Gets measured rich text segments on the line.
    /// </summary>
    public IReadOnlyList<OfficeRichTextSegment> Segments { get; }

    /// <summary>
    /// Gets the total measured line width.
    /// </summary>
    public double Width { get; }

    /// <summary>
    /// Gets the largest font size used by this line.
    /// </summary>
    public double FontSize { get; }

    /// <summary>
    /// Gets the resolved line height for this line, or zero when the containing block line height should be used.
    /// </summary>
    public double LineHeight { get; }

    /// <summary>
    /// Gets the additional x offset applied to this line inside the text frame.
    /// </summary>
    public double OffsetX { get; }
}

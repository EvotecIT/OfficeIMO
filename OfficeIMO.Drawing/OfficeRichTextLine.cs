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
    public OfficeRichTextLine(IReadOnlyList<OfficeRichTextSegment> segments) {
        Segments = segments ?? Array.Empty<OfficeRichTextSegment>();
        double width = 0D;
        double fontSize = 0D;
        for (int i = 0; i < Segments.Count; i++) {
            width += Segments[i].Width;
            fontSize = Math.Max(fontSize, Segments[i].FontSize);
        }

        Width = width;
        FontSize = fontSize;
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
}

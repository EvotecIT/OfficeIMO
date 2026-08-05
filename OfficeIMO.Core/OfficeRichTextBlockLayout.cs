using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes a measured bounded rich text block.
/// </summary>
public sealed class OfficeRichTextBlockLayout {
    /// <summary>
    /// Creates a measured rich text block.
    /// </summary>
    public OfficeRichTextBlockLayout(IReadOnlyList<OfficeRichTextLine> lines, double lineHeight, double width, double height, bool clipped = false) {
        Lines = lines ?? Array.Empty<OfficeRichTextLine>();
        LineHeight = lineHeight;
        Width = width;
        Height = height;
        Clipped = clipped;
    }

    /// <summary>
    /// Gets the visible measured lines.
    /// </summary>
    public IReadOnlyList<OfficeRichTextLine> Lines { get; }

    /// <summary>
    /// Gets the shared line height used to place lines in the block.
    /// </summary>
    public double LineHeight { get; }

    /// <summary>
    /// Gets the measured block width.
    /// </summary>
    public double Width { get; }

    /// <summary>
    /// Gets the measured block height.
    /// </summary>
    public double Height { get; }

    /// <summary>
    /// Gets whether the block was ellipsized or clipped.
    /// </summary>
    public bool Clipped { get; }
}

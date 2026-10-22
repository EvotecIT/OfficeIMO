namespace OfficeIMO.Pdf;

/// <summary>
/// Describes an explicit paragraph tab stop in PDF layout coordinates.
/// </summary>
public sealed class PdfTabStop {
    private double _position;
    private char? _decimalSeparator;
    // Generated navigation anchors its final tab to each actual flow frame.
    // Authored tab stops retain their ordinary absolute positions.
    internal bool IsFrameRightRelative { get; set; }

    /// <summary>
    /// Creates a paragraph tab stop at the specified position.
    /// </summary>
    /// <param name="position">The finite tab stop position, in points, relative to the paragraph text frame. Zero and negative positions are permitted.</param>
    /// <param name="alignment">The alignment applied to the text following the tab character.</param>
    /// <param name="leader">The leader drawn across the tab gap.</param>
    public PdfTabStop(double position, PdfTabAlignment alignment = PdfTabAlignment.Left, PdfTabLeaderStyle leader = PdfTabLeaderStyle.None) {
        Position = position;
        Alignment = alignment;
        Leader = leader;
    }

    /// <summary>Tab stop position, in points, relative to the paragraph text frame.</summary>
    /// <remarks>
    /// Zero is the main text-frame origin. Negative positions can be reached
    /// from a hanging first line. A stop is used only while ahead of the line cursor.
    /// </remarks>
    public double Position {
        get => _position;
        set {
            if (double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentException("Tab stop position must be a finite value.", nameof(Position));
            }

            _position = value;
        }
    }

    /// <summary>Alignment applied to the text following the tab character.</summary>
    public PdfTabAlignment Alignment { get; set; }

    /// <summary>Leader drawn across the tab gap.</summary>
    public PdfTabLeaderStyle Leader { get; set; }

    /// <summary>Gets or sets the character anchored by decimal alignment. Null accepts the first period or comma.</summary>
    /// <remarks>Grouping punctuation is ignored when a separator is specified. Text without that character is right-aligned.</remarks>
    public char? DecimalSeparator {
        get => _decimalSeparator;
        set {
            if (value.HasValue && (char.IsControl(value.Value) || char.IsSurrogate(value.Value)))
                throw new System.ArgumentException("A decimal tab separator must be one noncontrol Unicode character.", nameof(DecimalSeparator));
            _decimalSeparator = value;
        }
    }

    /// <summary>Creates a copy of this tab stop.</summary>
    public PdfTabStop Clone() => new PdfTabStop(Position, Alignment, Leader) {
        DecimalSeparator = DecimalSeparator, IsFrameRightRelative = IsFrameRightRelative
    };

    internal PdfTabStop ResolveFramePosition(double frameWidth) {
        PdfTabStop clone = Clone();
        if (IsFrameRightRelative) {
            clone.Position = System.Math.Max(0.001D, frameWidth - Position);
            clone.IsFrameRightRelative = false;
        }
        return clone;
    }
}

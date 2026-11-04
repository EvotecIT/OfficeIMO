namespace OfficeIMO.Pdf;

/// <summary>
/// Describes an explicit paragraph tab stop in PDF layout coordinates.
/// </summary>
public sealed class PdfTabStop {
    private global::OfficeIMO.OfficeLength _position;
    private char? _decimalSeparator;
    private PdfTabAlignment _alignment;
    private global::OfficeIMO.OfficeLength _barWidth = global::OfficeIMO.OfficeLength.Points(0.5D);
    // Generated navigation anchors its final tab to each actual flow frame.
    // Authored tab stops retain their ordinary absolute positions.
    internal bool IsFrameRightRelative { get; set; }

    /// <summary>
    /// Creates a paragraph tab stop at the specified position.
    /// </summary>
    /// <param name="position">The physical tab stop position relative to the paragraph text frame. Zero and negative positions are permitted.</param>
    /// <param name="alignment">Text alignment, or Bar for a stationary paragraph rule.</param>
    /// <param name="leader">The leader drawn across the tab gap.</param>
    public PdfTabStop(global::OfficeIMO.OfficeLength position, PdfTabAlignment alignment = PdfTabAlignment.Left, PdfTabLeaderStyle leader = PdfTabLeaderStyle.None) {
        Position = position;
        Alignment = alignment;
        Leader = leader;
    }

    /// <summary>Physical tab stop position relative to the paragraph text frame.</summary>
    /// <remarks>
    /// Zero is the main text-frame origin. Negative positions can be reached
    /// from a hanging first line. A text stop is used only while ahead of the line cursor.
    /// A bar rule is painted at its declared position independently of the text cursor.
    /// </remarks>
    public global::OfficeIMO.OfficeLength Position {
        get => _position;
        set {
            _position = value;
        }
    }

    /// <summary>Text alignment, or <see cref="PdfTabAlignment.Bar"/> for a stationary paragraph rule.</summary>
    public PdfTabAlignment Alignment {
        get => _alignment;
        set {
            if (value != PdfTabAlignment.Bar) Guard.TabAlignment(value, nameof(Alignment));
            _alignment = value;
        }
    }

    /// <summary>Positive physical stroke width for a bar rule. The default is 0.5pt.</summary>
    /// <remarks>Other tab alignments do not use this value. Bar rules have square caps.</remarks>
    public global::OfficeIMO.OfficeLength BarWidth {
        get => _barWidth;
        set {
            Guard.Positive(value.ToPoints(), nameof(BarWidth));
            _barWidth = value;
        }
    }

    /// <summary>Stroke color for a bar rule. The default is black; other alignments do not use it.</summary>
    public global::OfficeIMO.OfficeColor BarColor { get; set; } = global::OfficeIMO.OfficeColor.Black;

    /// <summary>Leader drawn across a text tab gap. Bar rules do not use a leader.</summary>
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
        DecimalSeparator = DecimalSeparator, IsFrameRightRelative = IsFrameRightRelative,
        BarWidth = BarWidth, BarColor = BarColor
    };

    internal PdfTabStop ResolveFramePosition(double frameWidth) {
        PdfTabStop clone = Clone();
        if (IsFrameRightRelative) {
            clone.Position = global::OfficeIMO.OfficeLength.Points(System.Math.Max(0.001D, frameWidth - Position.ToPoints()));
            clone.IsFrameRightRelative = false;
        }
        return clone;
    }
}

namespace OfficeIMO.Pdf;

/// <summary>
/// Defines an explicit paragraph tab stop in PDF layout points.
/// </summary>
public sealed class PdfTabStop {
    private double _position;

    /// <summary>Creates a paragraph tab stop.</summary>
    /// <param name="position">Tab stop position in points relative to the paragraph text frame.</param>
    /// <param name="alignment">Text alignment anchored at this tab stop.</param>
    /// <param name="leader">Leader fill rendered before the following text.</param>
    public PdfTabStop(double position, PdfTabAlignment alignment = PdfTabAlignment.Left, PdfTabLeaderStyle leader = PdfTabLeaderStyle.None) {
        Position = position;
        Alignment = alignment;
        Leader = leader;
    }

    /// <summary>Tab stop position in points relative to the paragraph text frame.</summary>
    public double Position {
        get => _position;
        set {
            Guard.NonNegative(value, nameof(value));
            _position = value;
        }
    }

    /// <summary>Text alignment anchored at this tab stop.</summary>
    public PdfTabAlignment Alignment {
        get => _alignment;
        set {
            Guard.TabAlignment(value, nameof(value));
            _alignment = value;
        }
    }

    /// <summary>Leader fill rendered before the following text.</summary>
    public PdfTabLeaderStyle Leader {
        get => _leader;
        set {
            Guard.TabLeaderStyle(value, nameof(value));
            _leader = value;
        }
    }

    private PdfTabAlignment _alignment;
    private PdfTabLeaderStyle _leader;

    /// <summary>Creates an independent copy of this tab stop.</summary>
    public PdfTabStop Clone() => new PdfTabStop(Position, Alignment, Leader);
}

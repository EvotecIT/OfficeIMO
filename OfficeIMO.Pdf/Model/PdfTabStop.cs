namespace OfficeIMO.Pdf;

/// <summary>
/// Describes an explicit paragraph tab stop in PDF layout coordinates.
/// </summary>
public sealed class PdfTabStop {
    private double _position;

    /// <summary>
    /// Creates a paragraph tab stop at the specified position.
    /// </summary>
    /// <param name="position">The tab stop position, in points, relative to the paragraph text frame.</param>
    /// <param name="alignment">The alignment applied to the text following the tab character.</param>
    /// <param name="leader">The leader drawn across the tab gap.</param>
    public PdfTabStop(double position, PdfTabAlignment alignment = PdfTabAlignment.Left, PdfTabLeaderStyle leader = PdfTabLeaderStyle.None) {
        Position = position;
        Alignment = alignment;
        Leader = leader;
    }

    /// <summary>Tab stop position, in points, relative to the paragraph text frame.</summary>
    public double Position {
        get => _position;
        set {
            if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentException("Tab stop position must be a positive finite value.", nameof(Position));
            }

            _position = value;
        }
    }

    /// <summary>Alignment applied to the text following the tab character.</summary>
    public PdfTabAlignment Alignment { get; set; }

    /// <summary>Leader drawn across the tab gap.</summary>
    public PdfTabLeaderStyle Leader { get; set; }

    /// <summary>Creates a copy of this tab stop.</summary>
    public PdfTabStop Clone() => new PdfTabStop(Position, Alignment, Leader);
}

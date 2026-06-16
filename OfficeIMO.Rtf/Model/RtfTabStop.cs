namespace OfficeIMO.Rtf;

/// <summary>
/// Paragraph tab stop in twips.
/// </summary>
public sealed class RtfTabStop {
    private int _positionTwips;

    /// <summary>Creates a tab stop at the specified position.</summary>
    public RtfTabStop(int positionTwips, RtfTabAlignment alignment = RtfTabAlignment.Left, RtfTabLeader leader = RtfTabLeader.None) {
        if (positionTwips < 0) throw new ArgumentOutOfRangeException(nameof(positionTwips), "Tab stop position cannot be negative.");
        PositionTwips = positionTwips;
        Alignment = alignment;
        Leader = leader;
    }

    /// <summary>Tab stop position in twips.</summary>
    public int PositionTwips {
        get => _positionTwips;
        set {
            if (value < 0) throw new ArgumentOutOfRangeException(nameof(value), "Tab stop position cannot be negative.");
            _positionTwips = value;
        }
    }

    /// <summary>Tab stop alignment.</summary>
    public RtfTabAlignment Alignment { get; set; }

    /// <summary>Tab leader style.</summary>
    public RtfTabLeader Leader { get; set; }

    /// <summary>Sets the tab stop alignment.</summary>
    public RtfTabStop SetAlignment(RtfTabAlignment alignment) {
        Alignment = alignment;
        return this;
    }

    /// <summary>Sets the tab leader style.</summary>
    public RtfTabStop SetLeader(RtfTabLeader leader) {
        Leader = leader;
        return this;
    }
}

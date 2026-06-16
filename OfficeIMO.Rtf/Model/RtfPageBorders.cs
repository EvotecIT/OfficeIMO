namespace OfficeIMO.Rtf;

/// <summary>
/// Page border settings for a document or section.
/// </summary>
public sealed class RtfPageBorders {
    /// <summary>Top page border.</summary>
    public RtfPageBorder Top { get; } = new RtfPageBorder();

    /// <summary>Bottom page border.</summary>
    public RtfPageBorder Bottom { get; } = new RtfPageBorder();

    /// <summary>Left page border.</summary>
    public RtfPageBorder Left { get; } = new RtfPageBorder();

    /// <summary>Right page border.</summary>
    public RtfPageBorder Right { get; } = new RtfPageBorder();

    /// <summary>Whether the page border surrounds the header area.</summary>
    public bool IncludeHeader { get; set; }

    /// <summary>Whether the page border surrounds the footer area.</summary>
    public bool IncludeFooter { get; set; }

    /// <summary>Whether page border edges should snap to page-border alignment.</summary>
    public bool SnapToPageBorder { get; set; }

    /// <summary>Page range to which the border applies.</summary>
    public RtfPageBorderScope? Scope { get; set; }

    /// <summary>Whether the page border is displayed behind document contents.</summary>
    public bool? DisplayBehindText { get; set; }

    /// <summary>Source edge used for border spacing.</summary>
    public RtfPageBorderOffset? OffsetFrom { get; set; }

    /// <summary>Sets page-border display options encoded by RTF <c>\pgbrdropt</c>.</summary>
    public RtfPageBorders SetDisplayOptions(
        RtfPageBorderScope? scope = null,
        bool? displayBehindText = null,
        RtfPageBorderOffset? offsetFrom = null) {
        Scope = scope;
        DisplayBehindText = displayBehindText;
        OffsetFrom = offsetFrom;
        return this;
    }

    /// <summary>Returns the page border for the requested side.</summary>
    public RtfPageBorder GetBorder(RtfPageBorderSide side) {
        switch (side) {
            case RtfPageBorderSide.Bottom:
                return Bottom;
            case RtfPageBorderSide.Left:
                return Left;
            case RtfPageBorderSide.Right:
                return Right;
            default:
                return Top;
        }
    }

    /// <summary>Whether any page-border setting is present.</summary>
    public bool HasAnyValue =>
        Top.HasAnyValue ||
        Bottom.HasAnyValue ||
        Left.HasAnyValue ||
        Right.HasAnyValue ||
        IncludeHeader ||
        IncludeFooter ||
        SnapToPageBorder ||
        Scope.HasValue ||
        DisplayBehindText.HasValue ||
        OffsetFrom.HasValue;

    internal void Clear() {
        Top.Clear();
        Bottom.Clear();
        Left.Clear();
        Right.Clear();
        IncludeHeader = false;
        IncludeFooter = false;
        SnapToPageBorder = false;
        Scope = null;
        DisplayBehindText = null;
        OffsetFrom = null;
    }
}

namespace OfficeIMO.Rtf;

/// <summary>
/// Absolute positioning and frame metadata for an RTF paragraph.
/// </summary>
public sealed class RtfParagraphFrame {
    /// <summary>Frame width in twips, represented by <c>\absw</c>.</summary>
    public int? WidthTwips { get; set; }

    /// <summary>Frame height in twips, represented by <c>\absh</c>. Negative values request exact height.</summary>
    public int? HeightTwips { get; set; }

    /// <summary>Horizontal reference frame.</summary>
    public RtfParagraphFrameHorizontalAnchor? HorizontalAnchor { get; set; }

    /// <summary>Vertical reference frame.</summary>
    public RtfParagraphFrameVerticalAnchor? VerticalAnchor { get; set; }

    /// <summary>Horizontal positioning mode.</summary>
    public RtfParagraphFrameHorizontalPosition? HorizontalPosition { get; set; }

    /// <summary>Absolute horizontal offset in twips used by <c>\posx</c> or <c>\posnegx</c>.</summary>
    public int? HorizontalPositionTwips { get; set; }

    /// <summary>Vertical positioning mode.</summary>
    public RtfParagraphFrameVerticalPosition? VerticalPosition { get; set; }

    /// <summary>Absolute vertical offset in twips used by <c>\posy</c> or <c>\posnegy</c>.</summary>
    public int? VerticalPositionTwips { get; set; }

    /// <summary>Whether the frame anchor is locked to its paragraph, represented by <c>\abslock</c>.</summary>
    public bool AnchorLocked { get; set; }

    /// <summary>Whether the frame avoids overlap, represented by <c>\absnoovrlp</c>.</summary>
    public bool? NoOverlap { get; set; }

    /// <summary>Whether main text should not wrap around the frame, represented by <c>\nowrap</c>.</summary>
    public bool NoWrap { get; set; }

    /// <summary>Distance in twips from surrounding text in all directions, represented by <c>\dxfrtext</c>.</summary>
    public int? TextWrapDistanceTwips { get; set; }

    /// <summary>Horizontal distance in twips from surrounding text, represented by <c>\dfrmtxtx</c>.</summary>
    public int? TextWrapDistanceHorizontalTwips { get; set; }

    /// <summary>Vertical distance in twips from surrounding text, represented by <c>\dfrmtxty</c>.</summary>
    public int? TextWrapDistanceVerticalTwips { get; set; }

    /// <summary>Whether main text flows underneath the frame, represented by <c>\overlay</c>.</summary>
    public bool OverlayText { get; set; }

    /// <summary>Number of lines occupied by a drop cap, represented by <c>\dropcapli</c>.</summary>
    public int? DropCapLines { get; set; }

    /// <summary>Drop-cap placement, represented by <c>\dropcapt</c>.</summary>
    public RtfDropCapKind? DropCapKind { get; set; }

    /// <summary>Whether any frame metadata has been set.</summary>
    public bool HasAnyValue =>
        WidthTwips.HasValue ||
        HeightTwips.HasValue ||
        HorizontalAnchor.HasValue ||
        VerticalAnchor.HasValue ||
        HorizontalPosition.HasValue ||
        HorizontalPositionTwips.HasValue ||
        VerticalPosition.HasValue ||
        VerticalPositionTwips.HasValue ||
        AnchorLocked ||
        NoOverlap.HasValue ||
        NoWrap ||
        TextWrapDistanceTwips.HasValue ||
        TextWrapDistanceHorizontalTwips.HasValue ||
        TextWrapDistanceVerticalTwips.HasValue ||
        OverlayText ||
        DropCapLines.HasValue ||
        DropCapKind.HasValue;

    /// <summary>Sets frame dimensions in twips.</summary>
    public RtfParagraphFrame SetSize(int? widthTwips = null, int? heightTwips = null) {
        WidthTwips = widthTwips;
        HeightTwips = heightTwips;
        return this;
    }

    /// <summary>Sets frame reference anchors.</summary>
    public RtfParagraphFrame SetAnchors(RtfParagraphFrameHorizontalAnchor? horizontalAnchor = null, RtfParagraphFrameVerticalAnchor? verticalAnchor = null) {
        HorizontalAnchor = horizontalAnchor;
        VerticalAnchor = verticalAnchor;
        return this;
    }

    /// <summary>Sets horizontal and vertical frame positioning.</summary>
    public RtfParagraphFrame SetPosition(
        RtfParagraphFrameHorizontalPosition? horizontalPosition = null,
        int? horizontalTwips = null,
        RtfParagraphFrameVerticalPosition? verticalPosition = null,
        int? verticalTwips = null) {
        HorizontalPosition = horizontalPosition;
        HorizontalPositionTwips = horizontalTwips;
        VerticalPosition = verticalPosition;
        VerticalPositionTwips = verticalTwips;
        return this;
    }

    /// <summary>Sets frame wrapping and overlay behavior.</summary>
    public RtfParagraphFrame SetWrapping(
        bool? noWrap = null,
        int? allDirectionsTwips = null,
        int? horizontalTwips = null,
        int? verticalTwips = null,
        bool? overlayText = null,
        bool? noOverlap = null) {
        if (noWrap.HasValue) {
            NoWrap = noWrap.Value;
        }

        TextWrapDistanceTwips = allDirectionsTwips;
        TextWrapDistanceHorizontalTwips = horizontalTwips;
        TextWrapDistanceVerticalTwips = verticalTwips;
        if (overlayText.HasValue) {
            OverlayText = overlayText.Value;
        }

        NoOverlap = noOverlap;
        return this;
    }

    /// <summary>Sets drop-cap metadata for this paragraph frame.</summary>
    public RtfParagraphFrame SetDropCap(int? lines, RtfDropCapKind? kind) {
        DropCapLines = lines;
        DropCapKind = kind;
        return this;
    }

    internal void Clear() {
        WidthTwips = null;
        HeightTwips = null;
        HorizontalAnchor = null;
        VerticalAnchor = null;
        HorizontalPosition = null;
        HorizontalPositionTwips = null;
        VerticalPosition = null;
        VerticalPositionTwips = null;
        AnchorLocked = false;
        NoOverlap = null;
        NoWrap = false;
        TextWrapDistanceTwips = null;
        TextWrapDistanceHorizontalTwips = null;
        TextWrapDistanceVerticalTwips = null;
        OverlayText = false;
        DropCapLines = null;
        DropCapKind = null;
    }

    internal void CopyFrom(RtfParagraphFrame source) {
        WidthTwips = source.WidthTwips;
        HeightTwips = source.HeightTwips;
        HorizontalAnchor = source.HorizontalAnchor;
        VerticalAnchor = source.VerticalAnchor;
        HorizontalPosition = source.HorizontalPosition;
        HorizontalPositionTwips = source.HorizontalPositionTwips;
        VerticalPosition = source.VerticalPosition;
        VerticalPositionTwips = source.VerticalPositionTwips;
        AnchorLocked = source.AnchorLocked;
        NoOverlap = source.NoOverlap;
        NoWrap = source.NoWrap;
        TextWrapDistanceTwips = source.TextWrapDistanceTwips;
        TextWrapDistanceHorizontalTwips = source.TextWrapDistanceHorizontalTwips;
        TextWrapDistanceVerticalTwips = source.TextWrapDistanceVerticalTwips;
        OverlayText = source.OverlayText;
        DropCapLines = source.DropCapLines;
        DropCapKind = source.DropCapKind;
    }
}

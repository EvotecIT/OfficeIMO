namespace OfficeIMO.Rtf.Html;

internal sealed class HtmlStyleDeclaration {
    internal static readonly HtmlStyleDeclaration Empty = new HtmlStyleDeclaration();

    internal bool? Bold { get; set; }

    internal bool? Italic { get; set; }

    internal bool? Underline { get; set; }

    internal RtfUnderlineStyle? UnderlineStyle { get; set; }

    internal RtfColor? UnderlineColor { get; set; }

    internal bool? Strike { get; set; }

    internal bool? DoubleStrike { get; set; }

    internal bool? Hidden { get; set; }

    internal bool? Outline { get; set; }

    internal bool? Shadow { get; set; }

    internal bool? Emboss { get; set; }

    internal bool? Imprint { get; set; }

    internal RtfCapsStyle? CapsStyle { get; set; }

    internal RtfVerticalPosition? VerticalPosition { get; set; }

    internal RtfTextDirection? Direction { get; set; }

    internal int? LanguageId { get; set; }

    internal RtfTextAlignment? TextAlignment { get; set; }

    internal RtfColor? ForegroundColor { get; set; }

    internal RtfColor? BackgroundColor { get; set; }

    internal string? FontFamily { get; set; }

    internal double? FontSizePoints { get; set; }

    internal int? CharacterSpacingTwips { get; set; }

    internal int? CharacterScalePercent { get; set; }

    internal int? CharacterOffsetHalfPoints { get; set; }

    internal int? LeftIndentTwips { get; set; }

    internal int? RightIndentTwips { get; set; }

    internal int? FirstLineIndentTwips { get; set; }

    internal int? SpaceBeforeTwips { get; set; }

    internal int? SpaceAfterTwips { get; set; }

    internal int? LineSpacingTwips { get; set; }

    internal bool? LineSpacingMultiple { get; set; }

    internal bool PageBreakBefore { get; set; }

    internal bool PageBreakAfter { get; set; }

    internal RtfTableCellVerticalAlignment? TableCellVerticalAlignment { get; set; }

    internal int? TableWidth { get; set; }

    internal RtfTableWidthUnit? TableWidthUnit { get; set; }

    internal int? TableHeightTwips { get; set; }

    internal int? PaddingTopTwips { get; set; }

    internal int? PaddingLeftTwips { get; set; }

    internal int? PaddingBottomTwips { get; set; }

    internal int? PaddingRightTwips { get; set; }

    internal HtmlBorderDeclaration? TopBorder { get; set; }

    internal HtmlBorderDeclaration? LeftBorder { get; set; }

    internal HtmlBorderDeclaration? BottomBorder { get; set; }

    internal HtmlBorderDeclaration? RightBorder { get; set; }

    internal bool? NoWrap { get; set; }

    internal bool HasBorderFormatting =>
        TopBorder != null ||
        LeftBorder != null ||
        BottomBorder != null ||
        RightBorder != null;

    internal bool HasInlineFormatting =>
        Bold.HasValue ||
        Italic.HasValue ||
        Underline.HasValue ||
        UnderlineStyle.HasValue ||
        UnderlineColor != null ||
        Strike.HasValue ||
        DoubleStrike.HasValue ||
        Hidden.HasValue ||
        Outline.HasValue ||
        Shadow.HasValue ||
        Emboss.HasValue ||
        Imprint.HasValue ||
        CapsStyle.HasValue ||
        VerticalPosition.HasValue ||
        Direction.HasValue ||
        LanguageId.HasValue ||
        ForegroundColor != null ||
        BackgroundColor != null ||
        FontFamily != null ||
        FontSizePoints.HasValue ||
        CharacterSpacingTwips.HasValue ||
        CharacterScalePercent.HasValue ||
        CharacterOffsetHalfPoints.HasValue ||
        HasBorderFormatting;
}

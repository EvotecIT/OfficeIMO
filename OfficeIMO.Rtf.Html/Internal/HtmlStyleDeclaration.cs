namespace OfficeIMO.Rtf.Html;

internal sealed class HtmlStyleDeclaration {
    internal static readonly HtmlStyleDeclaration Empty = new HtmlStyleDeclaration();

    internal bool? Bold { get; set; }

    internal bool? Italic { get; set; }

    internal bool? Underline { get; set; }

    internal bool? Strike { get; set; }

    internal RtfVerticalPosition? VerticalPosition { get; set; }

    internal RtfTextAlignment? TextAlignment { get; set; }

    internal RtfColor? ForegroundColor { get; set; }

    internal RtfColor? BackgroundColor { get; set; }

    internal string? FontFamily { get; set; }

    internal double? FontSizePoints { get; set; }

    internal int? LeftIndentTwips { get; set; }

    internal int? RightIndentTwips { get; set; }

    internal int? FirstLineIndentTwips { get; set; }

    internal bool PageBreakBefore { get; set; }

    internal bool PageBreakAfter { get; set; }

    internal bool HasInlineFormatting =>
        Bold.HasValue ||
        Italic.HasValue ||
        Underline.HasValue ||
        Strike.HasValue ||
        VerticalPosition.HasValue ||
        ForegroundColor != null ||
        BackgroundColor != null ||
        FontFamily != null ||
        FontSizePoints.HasValue;
}

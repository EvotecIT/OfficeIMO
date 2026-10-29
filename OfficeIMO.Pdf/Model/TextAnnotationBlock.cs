namespace OfficeIMO.Pdf;

internal sealed class TextAnnotationBlock : IPdfBlock {
    public TextAnnotationBlock(string contents, double width, double height, PdfAlign align, double spacingBefore, double spacingAfter, PdfTextAnnotationIcon icon, global::OfficeIMO.OfficeColor? color, bool open) {
        Contents = contents;
        Width = width;
        Height = height;
        Align = align;
        SpacingBefore = spacingBefore;
        SpacingAfter = spacingAfter;
        Icon = icon;
        Color = color;
        Open = open;
    }

    public string Contents { get; }
    public double Width { get; }
    public double Height { get; }
    public PdfAlign Align { get; }
    public double SpacingBefore { get; }
    public double SpacingAfter { get; }
    public PdfTextAnnotationIcon Icon { get; }
    public global::OfficeIMO.OfficeColor? Color { get; }
    public bool Open { get; }
}

internal sealed class FreeTextAnnotationBlock : IPdfBlock {
    public FreeTextAnnotationBlock(string contents, double width, double height, PdfAlign align, double spacingBefore, double spacingAfter, double fontSize, global::OfficeIMO.OfficeColor textColor, global::OfficeIMO.OfficeColor? borderColor, double borderWidth, global::OfficeIMO.OfficeColor? fillColor, PdfAlign textAlign, double padding, double? lineHeight) {
        Contents = contents;
        Width = width;
        Height = height;
        Align = align;
        SpacingBefore = spacingBefore;
        SpacingAfter = spacingAfter;
        FontSize = fontSize;
        TextColor = textColor;
        BorderColor = borderColor;
        BorderWidth = borderWidth;
        FillColor = fillColor;
        TextAlign = textAlign;
        Padding = padding;
        LineHeight = lineHeight;
    }

    public string Contents { get; }
    public double Width { get; }
    public double Height { get; }
    public PdfAlign Align { get; }
    public double SpacingBefore { get; }
    public double SpacingAfter { get; }
    public double FontSize { get; }
    public global::OfficeIMO.OfficeColor TextColor { get; }
    public global::OfficeIMO.OfficeColor? BorderColor { get; }
    public double BorderWidth { get; }
    public global::OfficeIMO.OfficeColor? FillColor { get; }
    public PdfAlign TextAlign { get; }
    public double Padding { get; }
    public double? LineHeight { get; }
}

internal sealed class HighlightAnnotationBlock : IPdfBlock {
    public HighlightAnnotationBlock(string contents, double width, double height, PdfAlign align, double spacingBefore, double spacingAfter, global::OfficeIMO.OfficeColor color) {
        Contents = contents;
        Width = width;
        Height = height;
        Align = align;
        SpacingBefore = spacingBefore;
        SpacingAfter = spacingAfter;
        Color = color;
    }

    public string Contents { get; }
    public double Width { get; }
    public double Height { get; }
    public PdfAlign Align { get; }
    public double SpacingBefore { get; }
    public double SpacingAfter { get; }
    public global::OfficeIMO.OfficeColor Color { get; }
}

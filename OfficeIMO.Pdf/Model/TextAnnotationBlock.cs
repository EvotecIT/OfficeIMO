namespace OfficeIMO.Pdf;

internal sealed class TextAnnotationBlock : IPdfBlock {
    public TextAnnotationBlock(string contents, double width, double height, PdfAlign align, double spacingBefore, double spacingAfter, PdfTextAnnotationIcon icon, PdfColor? color, bool open) {
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
    public PdfColor? Color { get; }
    public bool Open { get; }
}

internal sealed class FreeTextAnnotationBlock : IPdfBlock {
    public FreeTextAnnotationBlock(string contents, double width, double height, PdfAlign align, double spacingBefore, double spacingAfter, double fontSize, PdfColor textColor, PdfColor? borderColor, double borderWidth, PdfColor? fillColor, PdfAlign textAlign, double padding, double? lineHeight) {
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
    public PdfColor TextColor { get; }
    public PdfColor? BorderColor { get; }
    public double BorderWidth { get; }
    public PdfColor? FillColor { get; }
    public PdfAlign TextAlign { get; }
    public double Padding { get; }
    public double? LineHeight { get; }
}

internal sealed class HighlightAnnotationBlock : IPdfBlock {
    public HighlightAnnotationBlock(string contents, double width, double height, PdfAlign align, double spacingBefore, double spacingAfter, PdfColor color) {
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
    public PdfColor Color { get; }
}

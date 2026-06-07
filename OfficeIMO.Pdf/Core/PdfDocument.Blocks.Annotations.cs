namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Adds a PDF text annotation at the current flow position.</summary>
    public PdfDocument TextAnnotation(string contents, double width = 18D, double height = 18D, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfTextAnnotationIcon icon = PdfTextAnnotationIcon.Comment, PdfColor? color = null, bool open = false) {
        AddBlock(CreateTextAnnotationBlock(contents, width, height, align, spacingBefore, spacingAfter, icon, color, open));
        return this;
    }

    /// <summary>Adds a PDF free-text annotation at the current flow position.</summary>
    public PdfDocument FreeTextAnnotation(string contents, double width, double height, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, double fontSize = 10D, PdfColor? textColor = null, PdfColor? borderColor = null, double borderWidth = 1D, PdfColor? fillColor = null, PdfAlign textAlign = PdfAlign.Left, double padding = 3D, double? lineHeight = null) {
        AddBlock(CreateFreeTextAnnotationBlock(contents, width, height, align, spacingBefore, spacingAfter, fontSize, textColor, borderColor, borderWidth, fillColor, textAlign, padding, lineHeight));
        return this;
    }

    /// <summary>Adds a PDF highlight annotation rectangle at the current flow position.</summary>
    public PdfDocument HighlightAnnotation(string contents, double width, double height, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfColor? color = null) {
        AddBlock(CreateHighlightAnnotationBlock(contents, width, height, align, spacingBefore, spacingAfter, color));
        return this;
    }

    internal static TextAnnotationBlock CreateTextAnnotationBlock(string contents, double width, double height, PdfAlign? align, double? spacingBefore, double? spacingAfter, PdfTextAnnotationIcon icon, PdfColor? color, bool open) {
        Guard.NotNullOrWhiteSpace(contents, nameof(contents));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        PdfAlign resolvedAlign = align ?? PdfAlign.Left;
        Guard.LeftCenterRightAlign(resolvedAlign, nameof(align), "Text annotation");
        ValidateTextAnnotationIcon(icon, nameof(icon));
        double resolvedSpacingBefore = spacingBefore ?? 0D;
        double resolvedSpacingAfter = spacingAfter ?? 0D;
        Guard.NonNegative(resolvedSpacingBefore, nameof(spacingBefore));
        Guard.NonNegative(resolvedSpacingAfter, nameof(spacingAfter));
        return new TextAnnotationBlock(contents, width, height, resolvedAlign, resolvedSpacingBefore, resolvedSpacingAfter, icon, color, open);
    }

    internal static FreeTextAnnotationBlock CreateFreeTextAnnotationBlock(string contents, double width, double height, PdfAlign? align, double? spacingBefore, double? spacingAfter, double fontSize, PdfColor? textColor, PdfColor? borderColor, double borderWidth, PdfColor? fillColor, PdfAlign textAlign = PdfAlign.Left, double padding = 3D, double? lineHeight = null) {
        Guard.NotNullOrWhiteSpace(contents, nameof(contents));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.Positive(fontSize, nameof(fontSize));
        Guard.NonNegative(borderWidth, nameof(borderWidth));
        Guard.LeftCenterRightAlign(textAlign, nameof(textAlign), "Free text annotation text");
        Guard.NonNegative(padding, nameof(padding));
        if (lineHeight.HasValue) {
            Guard.Positive(lineHeight.Value, nameof(lineHeight));
        }

        PdfAlign resolvedAlign = align ?? PdfAlign.Left;
        Guard.LeftCenterRightAlign(resolvedAlign, nameof(align), "Free text annotation");
        double resolvedSpacingBefore = spacingBefore ?? 0D;
        double resolvedSpacingAfter = spacingAfter ?? 0D;
        Guard.NonNegative(resolvedSpacingBefore, nameof(spacingBefore));
        Guard.NonNegative(resolvedSpacingAfter, nameof(spacingAfter));
        return new FreeTextAnnotationBlock(contents, width, height, resolvedAlign, resolvedSpacingBefore, resolvedSpacingAfter, fontSize, textColor ?? PdfColor.Black, borderColor, borderWidth, fillColor, textAlign, padding, lineHeight);
    }

    internal static HighlightAnnotationBlock CreateHighlightAnnotationBlock(string contents, double width, double height, PdfAlign? align, double? spacingBefore, double? spacingAfter, PdfColor? color) {
        Guard.NotNullOrWhiteSpace(contents, nameof(contents));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        PdfAlign resolvedAlign = align ?? PdfAlign.Left;
        Guard.LeftCenterRightAlign(resolvedAlign, nameof(align), "Highlight annotation");
        double resolvedSpacingBefore = spacingBefore ?? 0D;
        double resolvedSpacingAfter = spacingAfter ?? 0D;
        Guard.NonNegative(resolvedSpacingBefore, nameof(spacingBefore));
        Guard.NonNegative(resolvedSpacingAfter, nameof(spacingAfter));
        return new HighlightAnnotationBlock(contents, width, height, resolvedAlign, resolvedSpacingBefore, resolvedSpacingAfter, color ?? new PdfColor(1D, 0.92D, 0.2D));
    }

    internal static void ValidateTextAnnotationIcon(PdfTextAnnotationIcon icon, string paramName) {
        if (icon != PdfTextAnnotationIcon.Comment &&
            icon != PdfTextAnnotationIcon.Key &&
            icon != PdfTextAnnotationIcon.Note &&
            icon != PdfTextAnnotationIcon.Help &&
            icon != PdfTextAnnotationIcon.NewParagraph &&
            icon != PdfTextAnnotationIcon.Paragraph &&
            icon != PdfTextAnnotationIcon.Insert) {
            throw new ArgumentOutOfRangeException(paramName, "PDF text annotation icon must be Comment, Key, Note, Help, NewParagraph, Paragraph, or Insert.");
        }
    }
}

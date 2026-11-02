namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Adds a PDF text annotation at the current flow position.</summary>
    internal PdfDocument TextAnnotation(string contents, double width = 18D, double height = 18D, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfTextAnnotationIcon icon = PdfTextAnnotationIcon.Comment, global::OfficeIMO.OfficeColor? color = null, bool open = false) {
        AddBlock(CreateTextAnnotationBlock(contents, width, height, align, spacingBefore, spacingAfter, icon, color, open));
        return this;
    }

    /// <summary>Adds a PDF free-text annotation at the current flow position.</summary>
    internal PdfDocument FreeTextAnnotation(string contents, double width, double height, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, double fontSize = 10D, global::OfficeIMO.OfficeColor? textColor = null, global::OfficeIMO.OfficeColor? borderColor = null, double borderWidth = 1D, global::OfficeIMO.OfficeColor? fillColor = null, PdfAlign textAlign = PdfAlign.Left, double padding = 3D, double? lineHeight = null) {
        AddBlock(CreateFreeTextAnnotationBlock(contents, width, height, align, spacingBefore, spacingAfter, fontSize, textColor, borderColor, borderWidth, fillColor, textAlign, padding, lineHeight));
        return this;
    }

    /// <summary>Adds a PDF highlight annotation rectangle at the current flow position.</summary>
    internal PdfDocument HighlightAnnotation(string contents, double width, double height, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, global::OfficeIMO.OfficeColor? color = null) {
        AddBlock(CreateHighlightAnnotationBlock(contents, width, height, align, spacingBefore, spacingAfter, color));
        return this;
    }

    internal static TextAnnotationBlock CreateTextAnnotationBlock(string contents, double width, double height, PdfAlign? align, double? spacingBefore, double? spacingAfter, PdfTextAnnotationIcon icon, global::OfficeIMO.OfficeColor? color, bool open) {
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

    internal static FreeTextAnnotationBlock CreateFreeTextAnnotationBlock(string contents, double width, double height, PdfAlign? align, double? spacingBefore, double? spacingAfter, double fontSize, global::OfficeIMO.OfficeColor? textColor, global::OfficeIMO.OfficeColor? borderColor, double borderWidth, global::OfficeIMO.OfficeColor? fillColor, PdfAlign textAlign = PdfAlign.Left, double padding = 3D, double? lineHeight = null) {
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
        return new FreeTextAnnotationBlock(contents, width, height, resolvedAlign, resolvedSpacingBefore, resolvedSpacingAfter, fontSize, textColor ?? global::OfficeIMO.OfficeColor.Black, borderColor, borderWidth, fillColor, textAlign, padding, lineHeight);
    }

    internal static HighlightAnnotationBlock CreateHighlightAnnotationBlock(string contents, double width, double height, PdfAlign? align, double? spacingBefore, double? spacingAfter, global::OfficeIMO.OfficeColor? color) {
        Guard.NotNullOrWhiteSpace(contents, nameof(contents));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        PdfAlign resolvedAlign = align ?? PdfAlign.Left;
        Guard.LeftCenterRightAlign(resolvedAlign, nameof(align), "Highlight annotation");
        double resolvedSpacingBefore = spacingBefore ?? 0D;
        double resolvedSpacingAfter = spacingAfter ?? 0D;
        Guard.NonNegative(resolvedSpacingBefore, nameof(spacingBefore));
        Guard.NonNegative(resolvedSpacingAfter, nameof(spacingAfter));
        return new HighlightAnnotationBlock(contents, width, height, resolvedAlign, resolvedSpacingBefore, resolvedSpacingAfter, color ?? global::OfficeIMO.OfficeColor.FromNormalizedRgb(1D, 0.92D, 0.2D));
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

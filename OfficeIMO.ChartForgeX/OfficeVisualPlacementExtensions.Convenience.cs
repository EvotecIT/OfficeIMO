using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.ChartForgeX;

public static partial class OfficeVisualPlacementExtensions {
    /// <summary>Inserts a visual in Word, constraining its natural width to the paragraph's content area when no size is supplied.</summary>
    /// <remarks>Use the overload with an out conversion result to inspect fidelity diagnostics.</remarks>
    public static WordImage AddVisualArtifact(this WordParagraph paragraph, VisualArtifact artifact,
        OfficeVisualConversionOptions? options = null, WordImageTextWrapping wrapping = WordImageTextWrapping.InLineWithText) =>
        paragraph.AddVisualArtifact(artifact, out _, options, wrapping);

    /// <summary>Inserts a visual anchored to a worksheet cell.</summary>
    /// <remarks>Use the overload with an out conversion result to inspect fidelity diagnostics.</remarks>
    public static ExcelImage AddVisualArtifact(this ExcelSheet sheet, int row, int column, VisualArtifact artifact,
        OfficeVisualConversionOptions? options = null, int offsetXPixels = 0, int offsetYPixels = 0) =>
        sheet.AddVisualArtifact(row, column, artifact, out _, options, offsetXPixels, offsetYPixels);

    /// <summary>Inserts a visual at a slide position expressed in points.</summary>
    /// <remarks>Use a layout box to constrain both dimensions, or the overload with an out result to inspect fidelity.</remarks>
    public static PowerPointPicture AddVisualArtifact(this PowerPointSlide slide, VisualArtifact artifact,
        double leftPoints = 0D, double topPoints = 0D, OfficeVisualConversionOptions? options = null) =>
        slide.AddVisualArtifact(artifact, leftPoints, topPoints, out _, options);

    /// <summary>Adds a visual to PDF flow, with proportional width constrained to the actual content frame by default.</summary>
    /// <remarks>Use the overload with an out conversion result to inspect fidelity diagnostics.</remarks>
    public static PdfContentBuilder AddVisualArtifact(this PdfContentBuilder content, VisualArtifact artifact,
        OfficeVisualConversionOptions? options = null, PdfAlign? align = null, double? spacingBefore = null,
        double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) =>
        content.AddVisualArtifact(artifact, out _, options, align, spacingBefore, spacingAfter, style, linkUri, linkContents);

    /// <summary>Adds a visual to an existing PDF document's flow.</summary>
    /// <remarks>Use the overload with an out conversion result to inspect fidelity diagnostics.</remarks>
    public static PdfDocument AddVisualArtifact(this PdfDocument document, VisualArtifact artifact,
        OfficeVisualConversionOptions? options = null, PdfAlign? align = null, double? spacingBefore = null,
        double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) =>
        document.AddVisualArtifact(artifact, out _, options, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
}

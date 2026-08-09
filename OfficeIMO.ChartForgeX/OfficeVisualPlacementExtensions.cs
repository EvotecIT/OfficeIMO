using System;
using System.IO;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.ChartForgeX;

/// <summary>Places converted ChartForgeX artifacts on OfficeIMO document surfaces.</summary>
public static class OfficeVisualPlacementExtensions {
    private const double PointsPerOfficePixel = 0.75D;

    /// <summary>Renders and inserts a ChartForgeX artifact as an SVG image in a Word paragraph.</summary>
    public static WordImage AddVisualArtifact(
        this WordParagraph paragraph,
        VisualArtifact artifact,
        out OfficeVisualConversionResult conversion,
        OfficeVisualConversionOptions? options = null,
        WordImageTextWrapping wrapping = WordImageTextWrapping.InLineWithText) {
        conversion = artifact.ToOfficeVisual(options);
        return paragraph.AddVisualArtifact(conversion, wrapping);
    }

    /// <summary>Inserts a converted ChartForgeX artifact as an SVG image in a Word paragraph.</summary>
    public static WordImage AddVisualArtifact(
        this WordParagraph paragraph,
        OfficeVisualConversionResult conversion,
        WordImageTextWrapping wrapping = WordImageTextWrapping.InLineWithText) {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        if (conversion == null) throw new ArgumentNullException(nameof(conversion));
        using var stream = new MemoryStream(conversion.GetPlacementBytes(), writable: false);
        WordImage image = paragraph.InsertImage(
            stream,
            ResolveFileName(conversion.Id, conversion.PlacementFileExtension),
            conversion.WidthPoints / PointsPerOfficePixel,
            conversion.HeightPoints / PointsPerOfficePixel,
            wrapping,
            conversion.AlternativeText);
        image.Title = ResolveTitle(conversion);
        return image;
    }

    /// <summary>Renders and inserts a ChartForgeX artifact as an SVG image anchored to an Excel cell.</summary>
    public static ExcelImage AddVisualArtifact(
        this ExcelSheet sheet,
        int row,
        int column,
        VisualArtifact artifact,
        out OfficeVisualConversionResult conversion,
        OfficeVisualConversionOptions? options = null,
        int offsetXPixels = 0,
        int offsetYPixels = 0) {
        conversion = artifact.ToOfficeVisual(options);
        return sheet.AddVisualArtifact(row, column, conversion, offsetXPixels, offsetYPixels);
    }

    /// <summary>Inserts a converted ChartForgeX artifact as an SVG image anchored to an Excel cell.</summary>
    public static ExcelImage AddVisualArtifact(
        this ExcelSheet sheet,
        int row,
        int column,
        OfficeVisualConversionResult conversion,
        int offsetXPixels = 0,
        int offsetYPixels = 0) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (conversion == null) throw new ArgumentNullException(nameof(conversion));
        return sheet.AddImage(
            row,
            column,
            conversion.GetPlacementBytes(),
            conversion.PlacementMediaType,
            ToPixels(conversion.WidthPoints),
            ToPixels(conversion.HeightPoints),
            offsetXPixels,
            offsetYPixels,
            ResolveTitle(conversion),
            conversion.AlternativeText,
            lockAspectRatio: true);
    }

    /// <summary>Renders and inserts a ChartForgeX artifact as an SVG image on a PowerPoint slide.</summary>
    public static PowerPointPicture AddVisualArtifact(
        this PowerPointSlide slide,
        VisualArtifact artifact,
        double leftPoints,
        double topPoints,
        out OfficeVisualConversionResult conversion,
        OfficeVisualConversionOptions? options = null) {
        conversion = artifact.ToOfficeVisual(options);
        return slide.AddVisualArtifact(conversion, leftPoints, topPoints);
    }

    /// <summary>Inserts a converted ChartForgeX artifact as an SVG image on a PowerPoint slide.</summary>
    public static PowerPointPicture AddVisualArtifact(
        this PowerPointSlide slide,
        OfficeVisualConversionResult conversion,
        double leftPoints = 0D,
        double topPoints = 0D) {
        if (slide == null) throw new ArgumentNullException(nameof(slide));
        if (conversion == null) throw new ArgumentNullException(nameof(conversion));
        ValidateNonNegativeFinite(leftPoints, nameof(leftPoints));
        ValidateNonNegativeFinite(topPoints, nameof(topPoints));
        using var stream = new MemoryStream(conversion.GetPlacementBytes(), writable: false);
        PowerPointPicture picture = slide.AddPicture(
            stream,
            conversion.PlacementFormat == OfficeVisualMediaFormat.Svg ? OfficeImageFormat.Svg : OfficeImageFormat.Png,
            PowerPointUnits.FromPoints(leftPoints),
            PowerPointUnits.FromPoints(topPoints),
            PowerPointUnits.FromPoints(conversion.WidthPoints),
            PowerPointUnits.FromPoints(conversion.HeightPoints));
        picture.Name = ResolveTitle(conversion);
        picture.Title = ResolveTitle(conversion);
        picture.Description = conversion.AlternativeText;
        return picture;
    }

    /// <summary>Renders and adds a ChartForgeX artifact to PDF flow content through the OfficeDrawing composition path.</summary>
    public static PdfContentCompose AddVisualArtifact(
        this PdfContentCompose content,
        VisualArtifact artifact,
        out OfficeVisualConversionResult conversion,
        OfficeVisualConversionOptions? options = null,
        PdfAlign? align = null,
        double? spacingBefore = null,
        double? spacingAfter = null,
        PdfDrawingStyle? style = null,
        string? linkUri = null,
        string? linkContents = null) {
        conversion = artifact.ToOfficeVisual(options);
        return content.AddVisualArtifact(conversion, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a converted ChartForgeX artifact to PDF flow content through the OfficeDrawing composition path.</summary>
    public static PdfContentCompose AddVisualArtifact(
        this PdfContentCompose content,
        OfficeVisualConversionResult conversion,
        PdfAlign? align = null,
        double? spacingBefore = null,
        double? spacingAfter = null,
        PdfDrawingStyle? style = null,
        string? linkUri = null,
        string? linkContents = null) {
        if (content == null) throw new ArgumentNullException(nameof(content));
        if (conversion == null) throw new ArgumentNullException(nameof(conversion));
        PdfDrawingStyle effectiveStyle = ResolvePdfDrawingStyle(style, conversion.AlternativeText);
        content.Item(item => item.Drawing(
            conversion.Drawing,
            align,
            spacingBefore,
            spacingAfter,
            effectiveStyle,
            linkUri,
            linkContents));
        return content;
    }

    /// <summary>Renders and adds a ChartForgeX artifact to top-level PDF flow through the OfficeDrawing composition path.</summary>
    public static PdfItemCompose AddVisualArtifact(
        this PdfItemCompose item,
        VisualArtifact artifact,
        out OfficeVisualConversionResult conversion,
        OfficeVisualConversionOptions? options = null,
        PdfAlign? align = null,
        double? spacingBefore = null,
        double? spacingAfter = null,
        PdfDrawingStyle? style = null,
        string? linkUri = null,
        string? linkContents = null) {
        conversion = artifact.ToOfficeVisual(options);
        return item.AddVisualArtifact(conversion, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a converted ChartForgeX artifact to top-level PDF flow through the OfficeDrawing composition path.</summary>
    public static PdfItemCompose AddVisualArtifact(
        this PdfItemCompose item,
        OfficeVisualConversionResult conversion,
        PdfAlign? align = null,
        double? spacingBefore = null,
        double? spacingAfter = null,
        PdfDrawingStyle? style = null,
        string? linkUri = null,
        string? linkContents = null) {
        if (item == null) throw new ArgumentNullException(nameof(item));
        if (conversion == null) throw new ArgumentNullException(nameof(conversion));
        return item.Drawing(
            conversion.Drawing,
            align,
            spacingBefore,
            spacingAfter,
            ResolvePdfDrawingStyle(style, conversion.AlternativeText),
            linkUri,
            linkContents);
    }

    /// <summary>Renders and adds a ChartForgeX artifact to an existing PDF document.</summary>
    public static PdfDocument AddVisualArtifact(
        this PdfDocument document,
        VisualArtifact artifact,
        out OfficeVisualConversionResult conversion,
        OfficeVisualConversionOptions? options = null,
        PdfAlign? align = null,
        double? spacingBefore = null,
        double? spacingAfter = null,
        PdfDrawingStyle? style = null,
        string? linkUri = null,
        string? linkContents = null) {
        conversion = artifact.ToOfficeVisual(options);
        return document.AddVisualArtifact(conversion, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a converted ChartForgeX artifact to an existing PDF document.</summary>
    public static PdfDocument AddVisualArtifact(
        this PdfDocument document,
        OfficeVisualConversionResult conversion,
        PdfAlign? align = null,
        double? spacingBefore = null,
        double? spacingAfter = null,
        PdfDrawingStyle? style = null,
        string? linkUri = null,
        string? linkContents = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (conversion == null) throw new ArgumentNullException(nameof(conversion));
        PdfDrawingStyle effectiveStyle = ResolvePdfDrawingStyle(style, conversion.AlternativeText);
        return document.Compose(compose => compose.Content(item => item.Drawing(
            conversion.Drawing,
            align,
            spacingBefore,
            spacingAfter,
            effectiveStyle,
            linkUri,
            linkContents)));
    }

    private static int ToPixels(double points) => Math.Max(1, (int)Math.Round(points / PointsPerOfficePixel, MidpointRounding.AwayFromZero));

    private static string ResolveTitle(OfficeVisualConversionResult conversion) =>
        string.IsNullOrWhiteSpace(conversion.Title) ? conversion.Id : conversion.Title;

    private static PdfDrawingStyle ResolvePdfDrawingStyle(PdfDrawingStyle? style, string alternativeText) {
        PdfDrawingStyle effective = style?.Clone() ?? new PdfDrawingStyle();
        if (!effective.Decorative && string.IsNullOrWhiteSpace(effective.AlternativeText)) {
            effective.AlternativeText = alternativeText;
        }

        return effective;
    }

    private static string ResolveFileName(string id, string extension) {
        string name = string.IsNullOrWhiteSpace(id) ? "chartforgex-visual" : id;
        foreach (char invalid in Path.GetInvalidFileNameChars()) name = name.Replace(invalid, '-');
        return name + extension;
    }

    private static void ValidateNonNegativeFinite(double value, string parameterName) {
        if (value < 0D || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, value, "Value must be non-negative and finite.");
        }
    }
}

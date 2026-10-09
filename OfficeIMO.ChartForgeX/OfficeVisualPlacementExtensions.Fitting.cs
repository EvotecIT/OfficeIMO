using System;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.ChartForgeX;

public static partial class OfficeVisualPlacementExtensions {
    /// <summary>Inserts a visual proportionally inside a slide layout box.</summary>
    public static PowerPointPicture AddVisualArtifact(this PowerPointSlide slide, VisualArtifact artifact,
        PowerPointLayoutBox box, OfficeVisualConversionOptions? options = null) =>
        slide.AddVisualArtifact(artifact, box, out _, options);

    /// <summary>Converts a visual and inserts it inside a slide layout box, returning the fitted conversion for inspection.</summary>
    /// <remarks>The box controls placement size. Options control conversion fidelity and proportional or stretched fitting.</remarks>
    public static PowerPointPicture AddVisualArtifact(this PowerPointSlide slide, VisualArtifact artifact,
        PowerPointLayoutBox box, out OfficeVisualConversionResult conversion, OfficeVisualConversionOptions? options = null) {
        if (slide == null) throw new ArgumentNullException(nameof(slide));
        conversion = artifact.ToOfficeVisual(options).WithSize(box.WidthPoints, box.HeightPoints, options?.Fit ?? OfficeImageFit.Contain);
        return PlaceInSlideBox(slide, conversion, box);
    }

    /// <summary>Places a reusable conversion inside a slide layout box without rendering it again.</summary>
    public static PowerPointPicture AddVisualArtifact(this PowerPointSlide slide, OfficeVisualConversionResult conversion,
        PowerPointLayoutBox box, OfficeImageFit fit = OfficeImageFit.Contain) {
        if (slide == null) throw new ArgumentNullException(nameof(slide));
        if (conversion == null) throw new ArgumentNullException(nameof(conversion));
        var fitted = conversion.WithSize(box.WidthPoints, box.HeightPoints, fit);
        return PlaceInSlideBox(slide, fitted, box);
    }

    /// <summary>Inserts a visual proportionally inside an authored worksheet range.</summary>
    public static ExcelImage AddVisualArtifact(this ExcelSheet sheet, string range, VisualArtifact artifact,
        OfficeVisualConversionOptions? options = null) => sheet.AddVisualArtifact(range, artifact, out _, options);

    /// <summary>Converts and inserts a visual inside a worksheet range, returning the fitted conversion for inspection.</summary>
    /// <remarks>Range dimensions control placement size. Options control conversion fidelity and proportional or stretched fitting.</remarks>
    public static ExcelImage AddVisualArtifact(this ExcelSheet sheet, string range, VisualArtifact artifact,
        out OfficeVisualConversionResult conversion, OfficeVisualConversionOptions? options = null) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        var bounds = sheet.GetRangeBounds(range);
        var size = sheet.GetRangeSizePixels(range);
        conversion = artifact.ToOfficeVisual(options).WithSize(size.WidthPixels * PointsPerOfficePixel,
            size.HeightPixels * PointsPerOfficePixel, options?.Fit ?? OfficeImageFit.Contain);
        return PlaceInRange(sheet, bounds.r1, bounds.c1, size.WidthPixels, size.HeightPixels, conversion);
    }

    /// <summary>Places a reusable conversion inside a worksheet range without rendering it again.</summary>
    public static ExcelImage AddVisualArtifact(this ExcelSheet sheet, string range, OfficeVisualConversionResult conversion,
        OfficeImageFit fit = OfficeImageFit.Contain) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (conversion == null) throw new ArgumentNullException(nameof(conversion));
        var bounds = sheet.GetRangeBounds(range);
        var size = sheet.GetRangeSizePixels(range);
        var fitted = conversion.WithSize(size.WidthPixels * PointsPerOfficePixel, size.HeightPixels * PointsPerOfficePixel, fit);
        return PlaceInRange(sheet, bounds.r1, bounds.c1, size.WidthPixels, size.HeightPixels, fitted);
    }

    private static PowerPointPicture PlaceInSlideBox(PowerPointSlide slide, OfficeVisualConversionResult fitted, PowerPointLayoutBox box) =>
        slide.AddVisualArtifact(fitted, box.LeftPoints + (box.WidthPoints - fitted.WidthPoints) / 2D,
            box.TopPoints + (box.HeightPoints - fitted.HeightPoints) / 2D);

    private static ExcelImage PlaceInRange(ExcelSheet sheet, int row, int column, int widthPixels, int heightPixels,
        OfficeVisualConversionResult fitted) {
        int offsetX = Math.Max(0, (int)Math.Round((widthPixels - fitted.WidthPoints / PointsPerOfficePixel) / 2D));
        int offsetY = Math.Max(0, (int)Math.Round((heightPixels - fitted.HeightPoints / PointsPerOfficePixel) / 2D));
        return sheet.AddVisualArtifact(row, column, fitted, offsetX, offsetY);
    }
}

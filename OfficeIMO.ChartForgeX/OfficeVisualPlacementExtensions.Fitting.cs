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
        conversion = artifact.ToOfficeVisual(options).WithSize(box.Width.ToPoints(), box.Height.ToPoints(), options?.Fit ?? OfficeImageFit.Contain);
        return PlaceInSlideBox(slide, conversion, box);
    }

    /// <summary>Places a reusable conversion inside a slide layout box without rendering it again.</summary>
    public static PowerPointPicture AddVisualArtifact(this PowerPointSlide slide, OfficeVisualConversionResult conversion,
        PowerPointLayoutBox box, OfficeImageFit fit = OfficeImageFit.Contain) {
        if (slide == null) throw new ArgumentNullException(nameof(slide));
        if (conversion == null) throw new ArgumentNullException(nameof(conversion));
        var fitted = conversion.WithSize(box.Width.ToPoints(), box.Height.ToPoints(), fit);
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
        var size = sheet.GetRangeSize(range);
        conversion = artifact.ToOfficeVisual(options).WithSize(size.Width.ToPoints(),
            size.Height.ToPoints(), options?.Fit ?? OfficeImageFit.Contain);
        return PlaceInRange(sheet, bounds.r1, bounds.c1, size.Width, size.Height, conversion);
    }

    /// <summary>Places a reusable conversion inside a worksheet range without rendering it again.</summary>
    public static ExcelImage AddVisualArtifact(this ExcelSheet sheet, string range, OfficeVisualConversionResult conversion,
        OfficeImageFit fit = OfficeImageFit.Contain) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (conversion == null) throw new ArgumentNullException(nameof(conversion));
        var bounds = sheet.GetRangeBounds(range);
        var size = sheet.GetRangeSize(range);
        var fitted = conversion.WithSize(size.Width.ToPoints(), size.Height.ToPoints(), fit);
        return PlaceInRange(sheet, bounds.r1, bounds.c1, size.Width, size.Height, fitted);
    }

    private static PowerPointPicture PlaceInSlideBox(PowerPointSlide slide, OfficeVisualConversionResult fitted, PowerPointLayoutBox box) =>
        slide.AddVisualArtifact(fitted, box.Left.ToPoints() + (box.Width.ToPoints() - fitted.WidthPoints) / 2D,
            box.Top.ToPoints() + (box.Height.ToPoints() - fitted.HeightPoints) / 2D);

    private static ExcelImage PlaceInRange(ExcelSheet sheet, int row, int column, OfficeLength width, OfficeLength height,
        OfficeVisualConversionResult fitted) {
        OfficeLength offsetX = OfficeLength.Emus(Math.Max(0L, (width - OfficeLength.Points(fitted.WidthPoints)).ToEmus() / 2));
        OfficeLength offsetY = OfficeLength.Emus(Math.Max(0L, (height - OfficeLength.Points(fitted.HeightPoints)).ToEmus() / 2));
        return sheet.AddVisualArtifact(row, column, fitted, offsetX, offsetY);
    }
}

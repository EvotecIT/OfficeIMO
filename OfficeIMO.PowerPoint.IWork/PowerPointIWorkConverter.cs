using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint.IWork;

namespace OfficeIMO.PowerPoint.IWork;

/// <summary>Projects Apple Keynote sources into editable OfficeIMO PowerPoint presentations.</summary>
public static partial class PowerPointIWorkConverter {
    private static KeynoteToPowerPointResult ProjectKeynote(
        IWorkSourceDocument source, IWorkConversionOptions? options = null) {
        CancellationToken cancellationToken = source.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        IWorkConversionOptions settings = (options ?? new IWorkConversionOptions()).Clone();
        IWorkConversionMode mode = settings.Mode;
        IWorkPreviewAsset? preview = mode == IWorkConversionMode.VisualOnly
            ? source.PreferredRasterPreview
            : null;
        if (mode == IWorkConversionMode.VisualOnly && preview == null) {
            throw new NotSupportedException("The Keynote source has no embedded raster preview.");
        }

        IWorkKeynoteProjection projection = source.ReadKeynote();
        string? destinationLimitation = mode == IWorkConversionMode.VisualOnly
            ? null
            : FindPowerPointProjectionLimitation(projection, settings.AllowPartialEditableReconstruction);
        bool hasEditableContent = projection.HasEditableContent
            || settings.AllowPartialEditableReconstruction && projection.HasRecoverableContent;
        bool editable = mode != IWorkConversionMode.VisualOnly && hasEditableContent
            && destinationLimitation == null;
        IReadOnlyList<IWorkDiagnostic> destinationDiagnostics =
            (!hasEditableContent || destinationLimitation == null
                ? Array.Empty<IWorkDiagnostic>()
                : new[] { new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_KEYNOTE_POWERPOINT_DESTINATION_UNSUPPORTED", destinationLimitation) })
            .Concat(editable
                ? FindPowerPointProjectionDiagnostics(projection, cancellationToken)
                : Array.Empty<IWorkDiagnostic>())
            .ToArray();
        if (editable && settings.AllowPartialEditableReconstruction &&
            (!projection.HasEditableContent || projection.Diagnostics.Any(diagnostic =>
                diagnostic.Severity != IWorkDiagnosticSeverity.Information)
                || destinationDiagnostics.Any(diagnostic => diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.Omission))) {
            destinationDiagnostics = destinationDiagnostics.Concat(new[] {
                new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_PARTIAL_EDITABLE_RECONSTRUCTION",
                    "Recovered editable content was retained under the explicit partial-reconstruction policy; source diagnostics describe incomplete details.")
            }).ToArray();
        }
        if (!editable && mode == IWorkConversionMode.EditableOnly) {
            throw new InvalidDataException(destinationLimitation
                ?? "The Keynote source has no supported editable slides.");
        }

        preview ??= editable ? null : source.PreferredRasterPreview;
        if (!editable && preview == null) {
            throw new NotSupportedException("The Keynote source has no supported editable slides or embedded raster preview.");
        }

        if (!editable) settings.ValidateVisualPreview(preview);

        PowerPointPresentation presentation = PowerPointPresentation.Create();
        try {
            IWorkCanvasSize? sourceSlideSize = projection.SlideSize;
            bool useSourceSlideSize = sourceSlideSize != null
                && FitsPresentationSlideMeasurement(sourceSlideSize.WidthPoints)
                && FitsPresentationSlideMeasurement(sourceSlideSize.HeightPoints);
            double canvasWidth = 960d;
            double canvasHeight = 540d;
            if (useSourceSlideSize && sourceSlideSize != null) {
                canvasWidth = sourceSlideSize.WidthPoints;
                canvasHeight = sourceSlideSize.HeightPoints;
                presentation.SlideSize.SetSizePoints(sourceSlideSize.WidthPoints,
                    sourceSlideSize.HeightPoints, PowerPointSlideSizeType.Custom);
            }
            if (editable) {
                var slidePairs = new List<(IWorkKeynoteSlide Source, PowerPointSlide Target)>(
                    projection.Slides.Count);
                foreach (IWorkKeynoteSlide sourceSlide in projection.Slides) {
                    cancellationToken.ThrowIfCancellationRequested();
                    PowerPointSlide slide = presentation.AddSlide();
                    if (sourceSlide.Name.Length > 0) slide.Name = sourceSlide.Name;
                    slide.Hidden = sourceSlide.IsSkipped;
                    slidePairs.Add((sourceSlide, slide));
                }
                foreach ((IWorkKeynoteSlide sourceSlide, PowerPointSlide slide) in slidePairs) {
                    cancellationToken.ThrowIfCancellationRequested();
                    foreach (IWorkKeynoteDrawable drawable in sourceSlide.Drawables) {
                        cancellationToken.ThrowIfCancellationRequested();
                        switch (drawable.Kind) {
                            case IWorkKeynoteDrawableKind.TextBox:
                                IWorkTextBox textBox = drawable.TextBox!;
                                if (drawable.IsTitlePlaceholder) {
                                    AddRichTextBox(slide, textBox,
                                        canvasWidth * 0.04875d, canvasHeight * 0.06d,
                                        canvasWidth * 0.9d, canvasHeight / 7.5d, cancellationToken);
                                } else {
                                    AddRichTextBox(slide, textBox,
                                        canvasWidth * 0.06375d, canvasHeight * 0.22d,
                                        canvasWidth * 0.87d, canvasHeight * 0.6533333333333333d, cancellationToken);
                                }
                                break;
                            case IWorkKeynoteDrawableKind.Table:
                                AddEditableTable(slide, drawable.Table!, cancellationToken);
                                break;
                            case IWorkKeynoteDrawableKind.Image:
                                AddEditableImage(slide, drawable.Image!, canvasWidth, canvasHeight);
                                break;
                        }
                    }
                    if (sourceSlide.PresenterNoteContent.Paragraphs.Count > 0) {
                        SetRichPresenterNotes(slide.Notes, sourceSlide.PresenterNoteContent, cancellationToken);
                    }
                }
            } else {
                PowerPointSlide slide = presentation.AddSlide();
                using var image = new MemoryStream(preview!.GetBytes(), writable: false);
                OfficeImageFormat format = preview.MediaType == "image/png"
                    ? OfficeImageFormat.Png
                    : OfficeImageFormat.Jpeg;
                (double left, double top, double width, double height) = PreviewLayout(preview,
                    canvasWidth / 72d, canvasHeight / 72d);
                PowerPointPicture picture = slide.AddPictureInches(image, format, left, top, width, height);
                picture.AltText = "Visual fallback from the source Keynote package";
            }

            IWorkProjectionKind kind = editable
                ? IWorkProjectionKind.EditableReconstruction
                : IWorkProjectionKind.VisualFallback;
            cancellationToken.ThrowIfCancellationRequested();
            return new KeynoteToPowerPointResult(presentation, source, projection,
                projection.CreateConversionReport(kind, preview, destinationDiagnostics,
                    settings.AllowPartialEditableReconstruction));
        } catch {
            presentation.Dispose();
            throw;
        }
    }

    private static (double Left, double Top, double Width, double Height) PreviewLayout(
        IWorkPreviewAsset preview, double slideWidth, double slideHeight) {
        double pixelWidth = preview.PixelWidth.GetValueOrDefault(16);
        double pixelHeight = preview.PixelHeight.GetValueOrDefault(9);
        double scale = Math.Min(slideWidth / pixelWidth, slideHeight / pixelHeight);
        double width = pixelWidth * scale;
        double height = pixelHeight * scale;
        return ((slideWidth - width) / 2d, (slideHeight - height) / 2d, width, height);
    }

    private static void AddEditableImage(PowerPointSlide slide, IWorkImageAsset source,
        double canvasWidth, double canvasHeight) {
        if (source.MediaType is not "image/png" and not "image/jpeg") return;
        double left = source.Geometry?.LeftPoints ?? 72;
        double top = source.Geometry?.TopPoints ?? 72;
        double width = source.Geometry?.WidthPoints
            ?? source.PixelWidth.GetValueOrDefault(640) * 72d / 96d;
        double height = source.Geometry?.HeightPoints
            ?? source.PixelHeight.GetValueOrDefault(480) * 72d / 96d;
        if (source.Geometry == null) {
            left = Math.Min(left, Math.Max(0d, canvasWidth - 1d));
            top = Math.Min(top, Math.Max(0d, canvasHeight - 1d));
            double scale = Math.Min(1d, Math.Min(
                Math.Max(1d, canvasWidth - left) / width,
                Math.Max(1d, canvasHeight - top) / height));
            width *= scale;
            height *= scale;
        }
        width = QuantizePositiveEmuPoints(width);
        height = QuantizePositiveEmuPoints(height);
        using var stream = new MemoryStream(source.GetBytes(), writable: false);
        PowerPointPicture picture = slide.AddPicturePoints(stream,
            source.MediaType == "image/png" ? OfficeImageFormat.Png : OfficeImageFormat.Jpeg,
            left, top, width, height);
        picture.Rotation = source.Geometry?.RotationDegrees;
        picture.AltText = source.AccessibilityDescription;
        if (source.Hyperlink != null
            && Uri.TryCreate(source.Hyperlink, UriKind.RelativeOrAbsolute, out Uri? imageLink)) {
            picture.SetHyperlink(imageLink);
        }
    }

    private static string? FindPowerPointProjectionLimitation(IWorkKeynoteProjection projection,
        bool allowPartialEditableReconstruction) {
        const double MaximumPointMeasurement = int.MaxValue / 12700d;
        const long MaximumDestinationTableCells = 1_000_000;
        long destinationTableCells = 0;
        if (projection.SlideSize is { } slideSize
            && (!FitsPresentationSlideMeasurement(slideSize.WidthPoints)
                || !FitsPresentationSlideMeasurement(slideSize.HeightPoints))) {
            return "The Keynote slide canvas has invalid dimensions.";
        }
        foreach (IWorkKeynoteSlide slide in projection.Slides) {
            IEnumerable<string?> drawableHyperlinks = slide.TextBoxes
                .Select(textBox => textBox.Hyperlink)
                .Concat(slide.TitleBox == null ? Array.Empty<string?>() : new[] { slide.TitleBox.Hyperlink })
                .Concat(slide.Images.Select(image => image.Hyperlink));
            if (drawableHyperlinks.Any(value => IsUnsupportedHyperlink(value, projection.Slides.Count))
                || SlideText(slide).SelectMany(content => content.Paragraphs)
                    .SelectMany(paragraph => paragraph.Runs)
                    .Select(run => run.Hyperlink)
                    .Concat(slide.Tables.SelectMany(table => table.Cells)
                        .Where(cell => cell.RichText != null)
                        .SelectMany(cell => cell.RichText!.Paragraphs)
                        .SelectMany(paragraph => paragraph.Runs)
                        .Select(run => run.Hyperlink))
                    .Any(value => IsUnsupportedHyperlink(value, projection.Slides.Count))) {
                return $"Keynote slide {slide.Index} contains a hyperlink that cannot be represented by the PPTX owner.";
            }
            IEnumerable<IWorkGeometry> geometries = slide.TextBoxes
                .Select(textBox => textBox.Geometry)
                .Concat(slide.TitleBox == null ? Array.Empty<IWorkGeometry?>() : new[] { slide.TitleBox.Geometry })
                .Concat(slide.Images.Select(image => image.Geometry))
                .Concat(slide.Tables.Select(table => table.Geometry))
                .Where(geometry => geometry != null)
                .Cast<IWorkGeometry>();
            if (geometries.Any(geometry => !IsFinite(geometry.LeftPoints)
                    || Math.Abs(geometry.LeftPoints) > MaximumPointMeasurement
                    || !IsFinite(geometry.TopPoints)
                    || Math.Abs(geometry.TopPoints) > MaximumPointMeasurement
                    || !FitsPositiveMeasurement(geometry.WidthPoints, MaximumPointMeasurement, allowZero: true)
                    || !FitsPositiveMeasurement(geometry.HeightPoints, MaximumPointMeasurement, allowZero: true))) {
                return $"Keynote slide {slide.Index} contains geometry outside the PPTX measurement range.";
            }
            IEnumerable<IWorkGeometry> rotatedShapes = slide.TextBoxes
                .Select(textBox => textBox.Geometry)
                .Concat(slide.TitleBox == null ? Array.Empty<IWorkGeometry?>() : new[] { slide.TitleBox.Geometry })
                .Concat(slide.Images.Select(image => image.Geometry))
                .Concat(slide.Tables.Select(table => table.Geometry))
                .Where(geometry => geometry != null)
                .Cast<IWorkGeometry>();
            if (rotatedShapes.Any(geometry => !FitsRotation(geometry.RotationDegrees))) {
                return $"Keynote slide {slide.Index} contains rotation outside the PPTX range.";
            }
            if (slide.Images.Any(image => image.Geometry is { } geometry
                    && (geometry.WidthPoints <= 0 || geometry.HeightPoints <= 0))) {
                return $"Keynote slide {slide.Index} contains a zero-sized image that cannot be represented by the PPTX image owner.";
            }
            foreach (IWorkTextContent content in SlideText(slide)) {
                foreach (IWorkTextParagraph paragraph in content.Paragraphs) {
                    if (paragraph.BreakKind is IWorkParagraphBreakKind.Section
                        or IWorkParagraphBreakKind.Layout or IWorkParagraphBreakKind.Page) {
                        return $"Keynote slide {slide.Index} contains a section, layout, or page break that cannot be represented inside PPTX slide text.";
                    }
                    if (paragraph.ListLevel > 8) {
                        return $"Keynote slide {slide.Index} contains a list nesting level outside the PPTX range.";
                    }
                    if (paragraph.ListLevel >= 0 && paragraph.ListLabel is { Length: > 0 } label
                        && (label.Length > 1 || label[0] is >= 'A' and <= 'Z' or >= 'a' and <= 'z')
                        && !TryParseNumbering(label, out _, out _)) {
                        return $"Keynote slide {slide.Index} contains a list marker that cannot be represented by native PPTX numbering.";
                    }
                    IWorkParagraphStyle style = paragraph.Style;
                    if (!allowPartialEditableReconstruction && (style.PageBreakBefore == true
                        || style.KeepWithNext == true || style.KeepLinesTogether == true)) {
                        return $"Keynote slide {slide.Index} contains paragraph pagination formatting that the PPTX owner cannot preserve.";
                    }
                    if (!FitsTextCoordinate(style.FirstLineIndentPoints)
                        || !FitsTextCoordinate(style.LeftIndentPoints)
                        || !FitsTextCoordinate(style.RightIndentPoints)
                        || Math.Abs(style.RightIndentPoints.GetValueOrDefault()) > 0.000001d
                        || !FitsSpacing(style.SpaceBeforePoints)
                        || !FitsSpacing(style.SpaceAfterPoints) || !FitsLineSpacing(style.LineSpacingMultiplier)
                        || style.TabStops != null && style.TabStops.Any(tab => !FitsTextCoordinate(tab.PositionPoints))) {
                        return $"Keynote slide {slide.Index} contains paragraph formatting outside the PPTX range.";
                    }
                    IEnumerable<IWorkTextStyle> textStyles = paragraph.Runs.Select(run => run.Style).Concat(new[] { style.TextStyle });
                    if (textStyles.Any(textStyle => textStyle.FontSizePoints is double fontSize
                            && (!IsFinite(fontSize) || fontSize < 1d || fontSize > 4000d
                                || fontSize * 100d != Math.Round(fontSize * 100d,
                                    MidpointRounding.AwayFromZero)))) {
                        return $"Keynote slide {slide.Index} contains a font size outside the PPTX range or hundredth-point precision.";
                    }
                    if (textStyles.Any(textStyle => textStyle.Color is { Alpha: < byte.MaxValue }
                            || textStyle.BackgroundColor is { Alpha: < byte.MaxValue })) {
                        return $"Keynote slide {slide.Index} contains transparent text colors that cannot be represented by the PPTX owner.";
                    }
                }
            }
            foreach (IWorkTable table in slide.Tables) {
                if (!allowPartialEditableReconstruction && table.Cells.Any(cell => cell.Comment != null))
                    return $"Keynote table '{table.Name}' contains cell comments that this PPTX projection cannot preserve.";
                if (!allowPartialEditableReconstruction && (table.HiddenRows.Count > 0 || table.HiddenColumns.Count > 0))
                    return $"Keynote table '{table.Name}' has hidden rows or columns that the PPTX table owner cannot preserve.";
                string? textStyleLimitation = FindTableTextStyleLimitation(table, allowPartialEditableReconstruction);
                if (textStyleLimitation != null) return textStyleLimitation;
                long tableCells = (long)table.RowCount * table.ColumnCount;
                if (table.RowCount == 0 || table.ColumnCount == 0) {
                    return $"Keynote table '{table.Name}' has no rows or columns and cannot be represented by the PPTX table owner.";
                }
                if (table.RowCount > 4096 || table.ColumnCount > 4096
                    || tableCells > 100_000) {
                    return $"Keynote table '{table.Name}' is too large for bounded PPTX table reconstruction.";
                }
                if (destinationTableCells > MaximumDestinationTableCells - tableCells) {
                    return "Keynote tables exceed the bounded PPTX destination cell budget.";
                }
                if (table.Cells.Any(cell => cell.Kind == IWorkCellKind.Formula
                    && (cell.Value == null || !cell.CachedValueIsComplete))) {
                    return $"Keynote table '{table.Name}' contains a formula without a complete cached value that the PPTX owner cannot evaluate.";
                }
                if (table.Cells.Any(cell => cell.Kind == IWorkCellKind.Formula
                    && cell.RichText is { IsComplete: false })) {
                    return $"Keynote table '{table.Name}' contains formula cached text with incomplete formatting that the PPTX owner cannot preserve.";
                }
                if (table.Cells.Any(cell => cell.Padding != null && PaddingPoints(cell.Padding)
                    .Any(points => Math.Round(points * 12700d) > int.MaxValue
                        || !allowPartialEditableReconstruction && !IsExactEmu(points)))) {
                    return $"Keynote table '{table.Name}' has cell padding outside the PPTX margin range or EMU precision.";
                }
                if (table.HasPopulatedCoveredMergeCells()) {
                    return $"Keynote table '{table.Name}' contains content in a covered merged cell that the PPTX owner cannot preserve.";
                }
                destinationTableCells += tableCells;
                double fallbackWidth = TableAxisExtent(table.ColumnCount, table.ColumnWidths,
                    table.DefaultColumnWidth, null, Math.Max(144d, 72d * table.ColumnCount));
                double fallbackHeight = TableAxisExtent(table.RowCount, table.RowHeights,
                    table.DefaultRowHeight, null, Math.Max(36d, 24d * table.RowCount));
                double projectedWidth = table.Geometry?.WidthPoints ?? fallbackWidth;
                double projectedHeight = table.Geometry?.HeightPoints ?? fallbackHeight;
                if (!FitsPositiveMeasurement(projectedWidth, MaximumPointMeasurement)
                    || !FitsPositiveMeasurement(projectedHeight, MaximumPointMeasurement)
                    || table.ColumnWidths.Count > 0 && !FitsPositiveMeasurement(
                        AxisSourceTotal(table.ColumnCount, table.ColumnWidths, table.DefaultColumnWidth, projectedWidth), double.MaxValue)
                    || table.RowHeights.Count > 0 && !FitsPositiveMeasurement(
                        AxisSourceTotal(table.RowCount, table.RowHeights, table.DefaultRowHeight, projectedHeight), double.MaxValue)) {
                    return $"Keynote table '{table.Name}' has sizing outside the PPTX measurement range.";
                }
                if (table.Geometry is { } geometry
                    && (geometry.WidthPoints <= 0 || geometry.HeightPoints <= 0)) {
                    return $"Keynote table '{table.Name}' has a zero-sized extent that cannot be represented by the PPTX table owner.";
                }
            }
        }
        return null;
    }

    private static bool FitsPositiveMeasurement(double value, double maximum, bool allowZero = false) =>
        !double.IsNaN(value) && !double.IsInfinity(value)
        && (allowZero ? value >= 0 : value > 0) && value <= maximum;

    private static bool FitsPresentationSlideMeasurement(double value) =>
        !double.IsNaN(value) && !double.IsInfinity(value)
        && value >= 914400d / 12700d
        && value <= 51206400d / 12700d;

    private static IReadOnlyList<IWorkDiagnostic> FindPowerPointProjectionDiagnostics(
        IWorkKeynoteProjection projection, CancellationToken cancellationToken) {
        bool requiresEmuRounding = projection.SlideSize is { } slideSize
                && (!IsExactEmu(slideSize.WidthPoints) || !IsExactEmu(slideSize.HeightPoints))
            || projection.Slides.SelectMany(slide => slide.TextBoxes
                    .Select(textBox => textBox.Geometry)
                    .Concat(slide.TitleBox == null
                        ? Array.Empty<IWorkGeometry?>()
                        : new[] { slide.TitleBox.Geometry })
                    .Concat(slide.Images.Select(image => image.Geometry))
                    .Concat(slide.Tables.Select(table => table.Geometry)))
                .Where(geometry => geometry != null)
                .Cast<IWorkGeometry>()
                .Any(geometry => !IsExactEmu(geometry.LeftPoints)
                    || !IsExactEmu(geometry.TopPoints)
                    || !IsExactEmu(geometry.WidthPoints)
                    || !IsExactEmu(geometry.HeightPoints))
            || projection.Slides.SelectMany(slide => slide.Tables)
                .Any(table => TableSizingRequiresEmuRounding(table))
            || projection.Slides.SelectMany(slide => slide.Tables).SelectMany(table => table.Cells)
                .Any(cell => cell.Padding != null && PaddingPoints(cell.Padding).Any(points => !IsExactEmu(points)));
        var diagnostics = new List<IWorkDiagnostic>();
        if (projection.Slides.SelectMany(slide => slide.Tables).Any(table => table.Cells.Any(cell => cell.Comment != null))) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_TABLE_COMMENTS_OMITTED",
                "PPTX table reconstruction omits native cell comments; qualified comments remain on the iWork projection.",
                lossKind: global::OfficeIMO.OfficeConversionLossKind.Omission));
        }
        if (projection.Slides.SelectMany(slide => slide.Tables).Any(table => table.HiddenRows.Count > 0 || table.HiddenColumns.Count > 0)) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_TABLE_VISIBILITY_OMITTED",
                "PPTX tables retain source-hidden rows and columns as visible content; their visibility is retained on the iWork projection.",
                lossKind: global::OfficeIMO.OfficeConversionLossKind.Omission));
        }
        if (projection.Slides.SelectMany(slide => slide.Tables).SelectMany(TableParagraphStyles)
            .Any(style => style.PageBreakBefore == true || style.KeepWithNext == true || style.KeepLinesTogether == true)) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_PARAGRAPH_PAGINATION_OMITTED",
                "PPTX retains supported text formatting but cannot represent source paragraph pagination flags; the iWork projection retains them.",
                lossKind: global::OfficeIMO.OfficeConversionLossKind.Omission));
        }
        diagnostics.AddRange(IWorkNumericDisplayDiagnostics.ForTextTables(
            projection.Slides.SelectMany(slide => slide.Tables).SelectMany(table => table.Cells), "KEYNOTE", "PPTX", cancellationToken));
        if (projection.Slides.SelectMany(slide => slide.Tables).Any(table =>
            AxisSizingIsScaled(table.ColumnCount, table.ColumnWidths, table.DefaultColumnWidth, table.Geometry?.WidthPoints)
            || AxisSizingIsScaled(table.RowCount, table.RowHeights, table.DefaultRowHeight, table.Geometry?.HeightPoints))) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_TABLE_SIZING_SCALED",
                "Individual Keynote table sizes were scaled proportionally to preserve the drawable extent.",
                lossKind: global::OfficeIMO.OfficeConversionLossKind.Approximation));
        }
        if (requiresEmuRounding) diagnostics.AddRange(new[] {
            new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_PPTX_PRECISION",
                "Keynote point measurements were quantized to the nearest PPTX EMU; the bounded source geometry remains available on the load result.")
        });
        return diagnostics;
    }

    private static bool TableSizingRequiresEmuRounding(IWorkTable table) {
        if (table.RowCount <= 0 || table.ColumnCount <= 0) return false;
        return AxisSizingRequiresEmuRounding(table.ColumnCount, table.ColumnWidths,
                table.DefaultColumnWidth, table.Geometry?.WidthPoints,
                Math.Max(144d, 72d * table.ColumnCount))
            || AxisSizingRequiresEmuRounding(table.RowCount, table.RowHeights,
                table.DefaultRowHeight, table.Geometry?.HeightPoints,
                Math.Max(36d, 24d * table.RowCount));
    }

    private static bool IsExactEmu(double points) {
        double scaled = points * PowerPointUnits.EmusPerPoint;
        return scaled == Math.Round(scaled)
            && PowerPointUnits.ToPoints(PowerPointUnits.FromPoints(points)) == points;
    }

    private static double QuantizePositiveEmuPoints(double points) =>
        PowerPointUnits.ToPoints(Math.Max(1L, PowerPointUnits.FromPoints(points)));

    private static IEnumerable<IWorkTextContent> SlideText(IWorkKeynoteSlide slide) {
        if (slide.TitleBox != null) yield return slide.TitleBox.Content;
        foreach (IWorkTextBox textBox in slide.TextBoxes) yield return textBox.Content;
        yield return slide.PresenterNoteContent;
        foreach (IWorkTable table in slide.Tables) {
            foreach (IWorkTableCell cell in table.Cells) {
                if (cell.RichText != null) yield return cell.RichText;
            }
        }
    }

    private static bool FitsTextCoordinate(double? points) {
        if (!points.HasValue) return true;
        double scaled = points.Value * PowerPointUnits.EmusPerPoint;
        return IsFinite(points.Value) && Math.Abs(points.Value) <= int.MaxValue / PowerPointUnits.EmusPerPoint
            && scaled == Math.Round(scaled);
    }

    private static bool FitsSpacing(double? points) => !points.HasValue
        || IsFinite(points.Value) && points.Value >= 0 && points.Value <= int.MaxValue / 100d
        && Math.Abs(points.Value - Math.Round(points.Value * 100d,
            MidpointRounding.AwayFromZero) / 100d) <= 0.00001d;

    private static bool FitsLineSpacing(double? multiplier) => !multiplier.HasValue
        || IsFinite(multiplier.Value) && multiplier.Value * 100000d >= 1d
        && multiplier.Value <= 132d
        && Math.Abs(multiplier.Value * 100000d - Math.Round(multiplier.Value * 100000d))
            <= Math.Max(1e-6d, multiplier.Value * 0.01d);

    private static bool FitsRotation(double degrees) {
        double scaled = degrees * 60000d;
        return IsFinite(degrees) && Math.Abs(degrees) <= int.MaxValue / 60000d
            && scaled == Math.Round(scaled);
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool IsUnsupportedHyperlink(string? value, int slideCount) {
        if (value == null) return false;
        if (!Uri.TryCreate(value, UriKind.RelativeOrAbsolute, out Uri? uri)) return true;
        return !uri.IsAbsoluteUri
            && (!PowerPointHyperlinkResolver.TryParseSlideFragment(uri, out int slideNumber)
                || slideNumber > slideCount);
    }

}

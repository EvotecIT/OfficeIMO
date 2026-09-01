using OfficeIMO.Drawing;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint.IWork;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointPresentation {
    /// <summary>Loads a Keynote source into the normal editable PowerPoint model, using a visual preview only when requested or necessary.</summary>
    public static PowerPointPresentation LoadKeynote(string path, IWorkReadOptions? options = null) =>
        LoadKeynoteWithReport(path, options).Document;

    /// <summary>Loads a Keynote stream into the normal editable PowerPoint model, using a visual preview only when requested or necessary.</summary>
    public static PowerPointPresentation LoadKeynote(Stream stream, IWorkReadOptions? options = null) =>
        LoadKeynoteWithReport(stream, options).Document;

    /// <summary>Loads a Keynote source and returns its PowerPoint projection, bounded source model, and loss report.</summary>
    public static IWorkKeynoteLoadResult LoadKeynoteWithReport(string path, IWorkReadOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return ProjectKeynote(IWorkSourceDocument.Open(path, IWorkDocumentKind.Keynote, options));
    }

    /// <summary>Loads a Keynote stream and returns its PowerPoint projection, bounded source model, and loss report.</summary>
    public static IWorkKeynoteLoadResult LoadKeynoteWithReport(Stream stream, IWorkReadOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        return ProjectKeynote(IWorkSourceDocument.Open(stream, IWorkDocumentKind.Keynote, options));
    }

    private static IWorkKeynoteLoadResult ProjectKeynote(IWorkSourceDocument source) {
        IWorkImportMode mode = source.RequestedImportMode;
        IWorkPreviewAsset? preview = mode == IWorkImportMode.VisualOnly
            ? source.PreferredRasterPreview
            : null;
        if (mode == IWorkImportMode.VisualOnly && preview == null) {
            throw new NotSupportedException("The Keynote source has no embedded raster preview.");
        }

        IWorkKeynoteProjection projection = source.ReadKeynote();
        string? destinationLimitation = mode == IWorkImportMode.VisualOnly
            ? null
            : FindPowerPointProjectionLimitation(projection);
        bool editable = mode != IWorkImportMode.VisualOnly && projection.HasEditableContent
            && destinationLimitation == null;
        IReadOnlyList<IWorkDiagnostic> destinationDiagnostics =
            (!projection.HasEditableContent || destinationLimitation == null
                ? Array.Empty<IWorkDiagnostic>()
                : new[] { new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                    "IWORK_KEYNOTE_POWERPOINT_DESTINATION_UNSUPPORTED", destinationLimitation) })
            .Concat(editable
                ? FindPowerPointProjectionDiagnostics(projection)
                : Array.Empty<IWorkDiagnostic>())
            .ToArray();
        if (!editable && mode == IWorkImportMode.EditableOnly) {
            throw new InvalidDataException(destinationLimitation
                ?? "The Keynote source has no supported editable slides.");
        }

        preview ??= editable ? null : source.PreferredRasterPreview;
        if (!editable && preview == null) {
            throw new NotSupportedException("The Keynote source has no supported editable slides or embedded raster preview.");
        }

        PowerPointPresentation presentation = Create();
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
                    PowerPointSlide slide = presentation.AddSlide();
                    if (sourceSlide.Name.Length > 0) slide.Name = sourceSlide.Name;
                    slide.Hidden = sourceSlide.IsSkipped;
                    slidePairs.Add((sourceSlide, slide));
                }
                foreach ((IWorkKeynoteSlide sourceSlide, PowerPointSlide slide) in slidePairs) {
                    foreach (IWorkKeynoteDrawable drawable in sourceSlide.Drawables) {
                        switch (drawable.Kind) {
                            case IWorkKeynoteDrawableKind.TextBox:
                                IWorkTextBox textBox = drawable.TextBox!;
                                if (drawable.IsTitlePlaceholder) {
                                    AddRichTextBox(slide, textBox,
                                        canvasWidth * 0.04875d, canvasHeight * 0.06d,
                                        canvasWidth * 0.9d, canvasHeight / 7.5d);
                                } else {
                                    AddRichTextBox(slide, textBox,
                                        canvasWidth * 0.06375d, canvasHeight * 0.22d,
                                        canvasWidth * 0.87d, canvasHeight * 0.6533333333333333d);
                                }
                                break;
                            case IWorkKeynoteDrawableKind.Table:
                                AddEditableTable(slide, drawable.Table!);
                                break;
                            case IWorkKeynoteDrawableKind.Image:
                                AddEditableImage(slide, drawable.Image!, canvasWidth, canvasHeight);
                                break;
                        }
                    }
                    if (sourceSlide.PresenterNoteContent.Paragraphs.Count > 0) {
                        SetRichPresenterNotes(slide.Notes, sourceSlide.PresenterNoteContent);
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
            return new IWorkKeynoteLoadResult(presentation, source, projection,
                projection.CreateImportReport(kind, preview, destinationDiagnostics));
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

    private static void AddEditableTable(PowerPointSlide slide, IWorkTable source) {
        if (source.RowCount == 0 || source.ColumnCount == 0) return;
        double left = source.Geometry?.LeftPoints ?? 72d;
        double top = source.Geometry?.TopPoints ?? 72d;
        double? columnWidth = source.DefaultColumnWidth is > 0
            ? QuantizePositiveEmuPoints(source.DefaultColumnWidth.Value)
            : null;
        double? rowHeight = source.DefaultRowHeight is > 0
            ? QuantizePositiveEmuPoints(source.DefaultRowHeight.Value)
            : null;
        double width = QuantizePositiveEmuPoints(source.Geometry is { WidthPoints: > 0 }
            ? source.Geometry.WidthPoints
            : columnWidth.HasValue
                ? columnWidth.Value * source.ColumnCount
                : Math.Max(144d, 72d * source.ColumnCount));
        double height = QuantizePositiveEmuPoints(source.Geometry is { HeightPoints: > 0 }
            ? source.Geometry.HeightPoints
            : rowHeight.HasValue
                ? rowHeight.Value * source.RowCount
                : Math.Max(36d, 24d * source.RowCount));
        PowerPointTable table = slide.AddTablePoints(source.RowCount, source.ColumnCount,
            left, top, width, height);
        table.AltText = source.AccessibilityDescription;
        table.Rotation = source.Geometry?.RotationDegrees ?? 0d;
        table.FirstRow = source.HeaderRowCount > 0;
        table.FirstColumn = source.HeaderColumnCount > 0;
        table.LastRow = source.FooterRowCount > 0;
        foreach (IWorkTableCell sourceCell in source.Cells) {
            PowerPointTableCell target = table.GetCell(sourceCell.Row - 1, sourceCell.Column - 1);
            target.Text = sourceCell.Kind == IWorkCellKind.Formula && sourceCell.Value != null
                ? sourceCell.CachedDisplayText
                : sourceCell.DisplayText;
            if (sourceCell.Row <= source.HeaderRowCount || sourceCell.Column <= source.HeaderColumnCount
                || sourceCell.Row > source.RowCount - source.FooterRowCount) target.Bold = true;
        }
        foreach (IWorkTableMergeRange merge in source.MergedRanges) {
            table.MergeCells(merge.FirstRow - 1, merge.FirstColumn - 1,
                merge.LastRow - 1, merge.LastColumn - 1);
        }
        if (columnWidth.HasValue && source.Geometry is not { WidthPoints: > 0 }) {
            for (int column = 0; column < source.ColumnCount; column++) {
                table.SetColumnWidthPoints(column, columnWidth.Value);
            }
        }
        if (rowHeight.HasValue && source.Geometry is not { HeightPoints: > 0 }) {
            for (int row = 0; row < source.RowCount; row++) {
                table.SetRowHeightPoints(row, rowHeight.Value);
            }
        }
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

    private static string? FindPowerPointProjectionLimitation(IWorkKeynoteProjection projection) {
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
                    if (style.PageBreakBefore == true || style.KeepWithNext == true
                        || style.KeepLinesTogether == true) {
                        return $"Keynote slide {slide.Index} contains paragraph pagination formatting that the PPTX owner cannot preserve.";
                    }
                    if (!FitsTextCoordinate(style.FirstLineIndentPoints)
                        || !FitsTextCoordinate(style.LeftIndentPoints)
                        || !FitsTextCoordinate(style.RightIndentPoints)
                        || Math.Abs(style.RightIndentPoints.GetValueOrDefault()) > 0.000001d
                        || !FitsSpacing(style.SpaceBeforePoints)
                        || !FitsSpacing(style.SpaceAfterPoints)) {
                        return $"Keynote slide {slide.Index} contains paragraph formatting outside the PPTX range.";
                    }
                    if (paragraph.Runs.Any(run => run.Style.FontSizePoints is double fontSize
                            && (!IsFinite(fontSize) || fontSize < 1d || fontSize > 4000d
                                || fontSize * 100d != Math.Round(fontSize * 100d,
                                    MidpointRounding.AwayFromZero)))) {
                        return $"Keynote slide {slide.Index} contains a font size outside the PPTX range or hundredth-point precision.";
                    }
                    if (paragraph.Runs.Any(run => run.Style.Color is { Alpha: < byte.MaxValue }
                            || run.Style.BackgroundColor is { Alpha: < byte.MaxValue })) {
                        return $"Keynote slide {slide.Index} contains transparent text colors that cannot be represented by the PPTX owner.";
                    }
                }
            }
            foreach (IWorkTable table in slide.Tables) {
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
                if (table.Cells.Any(cell => cell.Kind == IWorkCellKind.Formula && cell.Value == null)) {
                    return $"Keynote table '{table.Name}' contains an uncached formula that the PPTX owner cannot evaluate.";
                }
                if (projection.HasEditableContent && table.HasPopulatedCoveredMergeCells()) {
                    return $"Keynote table '{table.Name}' contains content in a covered merged cell that the PPTX owner cannot preserve.";
                }
                destinationTableCells += tableCells;
                double fallbackWidth = table.DefaultColumnWidth is > 0
                    ? table.DefaultColumnWidth.Value * table.ColumnCount
                    : Math.Max(144d, 72d * table.ColumnCount);
                double fallbackHeight = table.DefaultRowHeight is > 0
                    ? table.DefaultRowHeight.Value * table.RowCount
                    : Math.Max(36d, 24d * table.RowCount);
                double projectedWidth = table.Geometry?.WidthPoints ?? fallbackWidth;
                double projectedHeight = table.Geometry?.HeightPoints ?? fallbackHeight;
                if (!FitsPositiveMeasurement(projectedWidth, MaximumPointMeasurement)
                    || !FitsPositiveMeasurement(projectedHeight, MaximumPointMeasurement)) {
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
        IWorkKeynoteProjection projection) {
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
                .Any(table => TableSizingRequiresEmuRounding(table));
        if (!requiresEmuRounding) return Array.Empty<IWorkDiagnostic>();
        return new[] {
            new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_KEYNOTE_PPTX_PRECISION",
                "Keynote point measurements were quantized to the nearest PPTX EMU; the bounded source geometry remains available on the load result.")
        };
    }

    private static bool TableSizingRequiresEmuRounding(IWorkTable table) {
        if (table.RowCount <= 0 || table.ColumnCount <= 0) return false;
        return table.Geometry is not { WidthPoints: > 0 }
                && table.DefaultColumnWidth is > 0 && !IsExactEmu(table.DefaultColumnWidth.Value)
            || table.Geometry is not { HeightPoints: > 0 }
                && table.DefaultRowHeight is > 0 && !IsExactEmu(table.DefaultRowHeight.Value);
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

    private static void AddRichTextBox(PowerPointSlide slide, IWorkTextBox source,
        double fallbackLeft, double fallbackTop, double fallbackWidth, double fallbackHeight) {
        double left = source.Geometry?.LeftPoints ?? fallbackLeft;
        double top = source.Geometry?.TopPoints ?? fallbackTop;
        double width = source.Geometry?.WidthPoints ?? fallbackWidth;
        double height = source.Geometry?.HeightPoints ?? fallbackHeight;
        PowerPointTextBox textBox = slide.AddTextBoxPoints(string.Empty, left, top, width, height);
        textBox.Rotation = source.Geometry?.RotationDegrees;
        textBox.AltText = source.AccessibilityDescription;
        if (source.Hyperlink != null
            && Uri.TryCreate(source.Hyperlink, UriKind.RelativeOrAbsolute, out Uri? shapeLink)) {
            textBox.SetHyperlink(shapeLink);
        }
        textBox.Clear();
        bool first = true;
        var listState = new IWorkPowerPointListState();
        foreach (IWorkTextParagraph sourceParagraph in source.Content.Paragraphs) {
            PowerPointParagraph paragraph;
            if (first) {
                paragraph = textBox.Paragraphs[0];
                paragraph.Text = string.Empty;
                first = false;
            } else {
                paragraph = textBox.AddParagraph();
            }
            ApplyParagraphStyle(paragraph, sourceParagraph,
                listState.StartsAtSourceLabel(sourceParagraph));
            WriteParagraphContent(paragraph, sourceParagraph);
        }
    }

    private static void SetRichPresenterNotes(PowerPointNotes notes, IWorkTextContent source) {
        IReadOnlyList<PowerPointParagraph> paragraphs = notes.SetParagraphs(
            source.Paragraphs.Select(_ => string.Empty));
        var listState = new IWorkPowerPointListState();
        for (int paragraphIndex = 0; paragraphIndex < source.Paragraphs.Count; paragraphIndex++) {
            IWorkTextParagraph sourceParagraph = source.Paragraphs[paragraphIndex];
            PowerPointParagraph paragraph = paragraphs[paragraphIndex];
            ApplyParagraphStyle(paragraph, sourceParagraph,
                listState.StartsAtSourceLabel(sourceParagraph));
            WriteParagraphContent(paragraph, sourceParagraph);
        }
        notes.Save();
    }

    private static void ApplyParagraphStyle(PowerPointParagraph paragraph,
        IWorkTextParagraph source, bool startsAtSourceLabel) {
        IWorkParagraphStyle style = source.Style;
        bool rightToLeft = OfficeTextElements.ResolveBaseDirection(source.Text)
            == OfficeTextDirection.RightToLeft;
        paragraph.RightToLeft = rightToLeft;
        if (style.Alignment.HasValue) {
            paragraph.Alignment = style.Alignment.Value switch {
                IWorkTextAlignment.Natural => rightToLeft
                    ? PowerPointTextAlignment.Right
                    : PowerPointTextAlignment.Left,
                IWorkTextAlignment.Center => PowerPointTextAlignment.Center,
                IWorkTextAlignment.Right => PowerPointTextAlignment.Right,
                IWorkTextAlignment.Justified => PowerPointTextAlignment.Justified,
                _ => PowerPointTextAlignment.Left
            };
        }
        paragraph.IndentPoints = style.FirstLineIndentPoints;
        paragraph.LeftMarginPoints = style.LeftIndentPoints;
        paragraph.SpaceBeforePoints = style.SpaceBeforePoints;
        paragraph.SpaceAfterPoints = style.SpaceAfterPoints;
        if (source.ListLevel >= 0) {
            paragraph.Level = Math.Min(8, source.ListLevel);
            if (string.IsNullOrEmpty(source.ListLabel)) paragraph.SetBullet('\u2022');
            else if (TryParseNumbering(source.ListLabel!, out PowerPointNumberingScheme scheme,
                         out int start)) {
                if (startsAtSourceLabel) paragraph.SetNumbered(scheme, start);
                else paragraph.SetNumbered(scheme);
            } else if (source.ListLabel!.Length == 1) paragraph.SetBullet(source.ListLabel[0]);
        }
    }

    private sealed class IWorkPowerPointListState {
        private readonly HashSet<int> _observedLevels = new();
        private bool _inList;
        private ulong? _listIdentifier;

        internal bool StartsAtSourceLabel(IWorkTextParagraph paragraph) {
            if (paragraph.ListLevel < 0) {
                _inList = false;
                _listIdentifier = null;
                _observedLevels.Clear();
                return false;
            }
            if (!_inList || paragraph.ListIdentifier != _listIdentifier) {
                _inList = true;
                _listIdentifier = paragraph.ListIdentifier;
                _observedLevels.Clear();
            }
            return _observedLevels.Add(paragraph.ListLevel);
        }
    }

    private static void WriteParagraphContent(PowerPointParagraph paragraph,
        IWorkTextParagraph source) {
        paragraph.Text = string.Empty;
        bool canReuseInitialRun = true;
        foreach (IWorkTextRun sourceRun in source.Runs) {
            string[] lines = sourceRun.Text.Split(new[] { '\n' });
            for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++) {
                if (lineIndex > 0) {
                    paragraph.AddLineBreak();
                    canReuseInitialRun = false;
                }
                if (lines[lineIndex].Length > 0) {
                    AppendStyledText(paragraph, lines[lineIndex], sourceRun.Style,
                        sourceRun.Hyperlink, ref canReuseInitialRun);
                }
            }
        }
    }

    private static bool TryParseNumbering(string label,
        out PowerPointNumberingScheme scheme, out int start) {
        const int MaximumStart = 32_767;
        scheme = PowerPointNumberingScheme.ArabicPeriod;
        start = 1;
        string marker = label.Trim();
        bool parenthesized = marker.Length > 2
            && marker[0] == '(' && marker[marker.Length - 1] == ')';
        bool rightParenthesis = !parenthesized && marker.EndsWith(")", StringComparison.Ordinal);
        bool period = !parenthesized && marker.EndsWith(".", StringComparison.Ordinal);
        string token = parenthesized
            ? marker.Substring(1, marker.Length - 2)
            : rightParenthesis || period ? marker.Substring(0, marker.Length - 1) : marker;
        if (token.Length == 0) return false;

        if (token.All(character => character is >= '0' and <= '9')) {
            if (!int.TryParse(token, System.Globalization.NumberStyles.None,
                    System.Globalization.CultureInfo.InvariantCulture, out start)
                || start is < 1 or > MaximumStart) return false;
            scheme = parenthesized ? PowerPointNumberingScheme.ArabicParenBoth
                : rightParenthesis ? PowerPointNumberingScheme.ArabicParenR
                : period ? PowerPointNumberingScheme.ArabicPeriod
                : PowerPointNumberingScheme.ArabicPlain;
            return true;
        }

        bool roman = token.All(character => "ivxlcdmIVXLCDM".IndexOf(character) >= 0)
            && (token.Length > 1 || "ivxIVX".IndexOf(token[0]) >= 0);
        if (roman) {
            if (!TryParseRoman(token, out start) || start > MaximumStart) return false;
            bool upper = token.All(character => character is >= 'A' and <= 'Z');
            if (!parenthesized && !rightParenthesis && !period) return false;
            scheme = upper
                ? parenthesized ? PowerPointNumberingScheme.RomanUpperCharacterParenBoth
                    : rightParenthesis ? PowerPointNumberingScheme.RomanUpperCharacterParenR
                    : PowerPointNumberingScheme.RomanUpperCharacterPeriod
                : parenthesized ? PowerPointNumberingScheme.RomanLowerCharacterParenBoth
                    : rightParenthesis ? PowerPointNumberingScheme.RomanLowerCharacterParenR
                    : PowerPointNumberingScheme.RomanLowerCharacterPeriod;
            return true;
        }

        bool uppercase = token.All(character => character is >= 'A' and <= 'Z');
        bool lowercase = token.All(character => character is >= 'a' and <= 'z');
        if (!uppercase && !lowercase
            || !TryParseAlphabetic(token, out start) || start > MaximumStart
            || !parenthesized && !rightParenthesis && !period) return false;
        scheme = uppercase
            ? parenthesized ? PowerPointNumberingScheme.AlphaUpperCharacterParenBoth
                : rightParenthesis ? PowerPointNumberingScheme.AlphaUpperCharacterParenR
                : PowerPointNumberingScheme.AlphaUpperCharacterPeriod
            : parenthesized ? PowerPointNumberingScheme.AlphaLowerCharacterParenBoth
                : rightParenthesis ? PowerPointNumberingScheme.AlphaLowerCharacterParenR
                : PowerPointNumberingScheme.AlphaLowerCharacterPeriod;
        return true;
    }

    private static bool TryParseAlphabetic(string token, out int value) {
        value = 0;
        foreach (char character in token) {
            int digit = char.ToUpperInvariant(character) - 'A' + 1;
            if (digit < 1 || digit > 26 || value > (int.MaxValue - digit) / 26) return false;
            value = value * 26 + digit;
        }
        return value > 0;
    }

    private static bool TryParseRoman(string token, out int value) {
        value = 0;
        int previous = 0;
        for (int index = token.Length - 1; index >= 0; index--) {
            int current = char.ToUpperInvariant(token[index]) switch {
                'I' => 1, 'V' => 5, 'X' => 10, 'L' => 50,
                'C' => 100, 'D' => 500, 'M' => 1000, _ => 0
            };
            if (current == 0) return false;
            int delta = current < previous ? -current : current;
            if (delta > 0 && value > int.MaxValue - delta
                || delta < 0 && value < int.MinValue - delta) return false;
            value += delta;
            if (current > previous) previous = current;
        }
        if (value <= 0) return false;
        bool upper = token.All(character => character is >= 'A' and <= 'Z');
        string canonical = FormatRoman(value);
        return string.Equals(token, upper ? canonical : canonical.ToLowerInvariant(),
            StringComparison.Ordinal);
    }

    private static string FormatRoman(int value) {
        var builder = new System.Text.StringBuilder();
        foreach ((int Number, string Token) part in new[] {
                     (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                     (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                     (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
                 }) {
            while (value >= part.Number) {
                builder.Append(part.Token);
                value -= part.Number;
            }
        }
        return builder.ToString();
    }

    private static void AppendStyledText(PowerPointParagraph paragraph, string text,
        IWorkTextStyle style, string? hyperlink, ref bool canReuseInitialRun) {
        PowerPointTextRun run;
        if (canReuseInitialRun) {
            run = paragraph.Runs[0];
            run.Text = text;
            canReuseInitialRun = false;
        } else {
            run = paragraph.AddRun(text);
        }
        ApplyTextStyle(run, style);
        if (hyperlink != null
            && Uri.TryCreate(hyperlink, UriKind.RelativeOrAbsolute, out Uri? runLink)) {
            run.Hyperlink = runLink;
        }
    }

    private static void ApplyTextStyle(PowerPointTextRun run, IWorkTextStyle style) {
        if (style.Bold.HasValue) run.Bold = style.Bold.Value;
        if (style.Italic.HasValue) run.Italic = style.Italic.Value;
        if (style.Underline.HasValue) run.Underline = style.Underline.Value;
        if (style.Strikethrough.HasValue) run.Strikethrough = style.Strikethrough.Value;
        if (style.FontSizePoints.HasValue) run.FontSizePoints = style.FontSizePoints.Value;
        if (!string.IsNullOrWhiteSpace(style.FontName)) run.FontName = style.FontName;
        if (style.Color != null) run.Color = style.Color.RgbHex;
        if (style.BackgroundColor != null) run.HighlightColor = style.BackgroundColor.RgbHex;
    }
}

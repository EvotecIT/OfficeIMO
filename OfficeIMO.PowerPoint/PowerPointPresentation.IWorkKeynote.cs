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
            if (editable) {
                if (projection.SlideSize != null) {
                    presentation.SlideSize.SetSizePoints(projection.SlideSize.WidthPoints,
                        projection.SlideSize.HeightPoints, PowerPointSlideSizeType.Custom);
                }
                foreach (IWorkKeynoteSlide sourceSlide in projection.Slides) {
                    PowerPointSlide slide = presentation.AddSlide();
                    if (sourceSlide.Name.Length > 0) slide.Name = sourceSlide.Name;
                    slide.Hidden = sourceSlide.IsSkipped;
                    if (sourceSlide.TitleBox != null) {
                        AddRichTextBox(slide, sourceSlide.TitleBox,
                            46.8, 32.4, 864, 72);
                    }
                    foreach (IWorkTextBox textBox in sourceSlide.TextBoxes) {
                        AddRichTextBox(slide, textBox,
                            61.2, 118.8, 835.2, 352.8);
                    }
                    foreach (IWorkTable sourceTable in sourceSlide.Tables) {
                        AddEditableTable(slide, sourceTable);
                    }
                    foreach (IWorkImageAsset sourceImage in sourceSlide.Images.Where(image =>
                                 image.MediaType is "image/png" or "image/jpeg")) {
                        double left = sourceImage.Geometry?.LeftPoints ?? 72;
                        double top = sourceImage.Geometry?.TopPoints ?? 72;
                        double width = sourceImage.Geometry?.WidthPoints
                            ?? sourceImage.PixelWidth.GetValueOrDefault(640) * 72d / 96d;
                        double height = sourceImage.Geometry?.HeightPoints
                            ?? sourceImage.PixelHeight.GetValueOrDefault(480) * 72d / 96d;
                        using var stream = new MemoryStream(sourceImage.GetBytes(), writable: false);
                        PowerPointPicture picture = slide.AddPicturePoints(stream,
                            sourceImage.MediaType == "image/png" ? OfficeImageFormat.Png : OfficeImageFormat.Jpeg,
                            left, top, width, height);
                        picture.Rotation = sourceImage.Geometry?.RotationDegrees;
                        picture.AltText = sourceImage.AccessibilityDescription;
                        if (sourceImage.Hyperlink != null
                            && Uri.TryCreate(sourceImage.Hyperlink, UriKind.Absolute,
                                out Uri? imageLink)) {
                            picture.SetHyperlink(imageLink);
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
                (double left, double top, double width, double height) = PreviewLayout(preview);
                slide.AddPictureInches(image, format, left, top, width, height);
            }

            IWorkProjectionKind kind = editable
                ? IWorkProjectionKind.EditableReconstruction
                : IWorkProjectionKind.VisualFallback;
            return new IWorkKeynoteLoadResult(presentation, source, projection, projection.CreateImportReport(kind, preview));
        } catch {
            presentation.Dispose();
            throw;
        }
    }

    private static (double Left, double Top, double Width, double Height) PreviewLayout(IWorkPreviewAsset preview) {
        const double slideWidth = 13.333;
        const double slideHeight = 7.5;
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
        double width = source.Geometry is { WidthPoints: > 0 }
            ? source.Geometry.WidthPoints
            : Math.Max(144d, source.DefaultColumnWidth.GetValueOrDefault(72d) * source.ColumnCount);
        double height = source.Geometry is { HeightPoints: > 0 }
            ? source.Geometry.HeightPoints
            : Math.Max(36d, source.DefaultRowHeight.GetValueOrDefault(24d) * source.RowCount);
        PowerPointTable table = slide.AddTablePoints(source.RowCount, source.ColumnCount,
            left, top, width, height);
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
        if (source.DefaultColumnWidth is > 0) {
            double columnWidth = width / source.ColumnCount;
            for (int column = 0; column < source.ColumnCount; column++) {
                table.SetColumnWidthPoints(column, columnWidth);
            }
        }
        if (source.DefaultRowHeight is > 0) {
            double rowHeight = height / source.RowCount;
            for (int row = 0; row < source.RowCount; row++) {
                table.SetRowHeightPoints(row, rowHeight);
            }
        }
    }

    private static string? FindPowerPointProjectionLimitation(IWorkKeynoteProjection projection) {
        const double MaximumPointMeasurement = int.MaxValue / 12700d;
        const long MaximumDestinationTableCells = 1_000_000;
        long destinationTableCells = 0;
        if (projection.SlideSize is { } slideSize
            && (!FitsPositiveMeasurement(slideSize.WidthPoints, MaximumPointMeasurement)
                || !FitsPositiveMeasurement(slideSize.HeightPoints, MaximumPointMeasurement))) {
            return "The Keynote slide canvas has invalid dimensions.";
        }
        foreach (IWorkKeynoteSlide slide in projection.Slides) {
            IEnumerable<string?> drawableHyperlinks = slide.TextBoxes
                .Select(textBox => textBox.Hyperlink)
                .Concat(slide.TitleBox == null ? Array.Empty<string?>() : new[] { slide.TitleBox.Hyperlink })
                .Concat(slide.Images.Select(image => image.Hyperlink));
            if (drawableHyperlinks.Any(IsUnsupportedExternalHyperlink)
                || SlideText(slide).SelectMany(content => content.Paragraphs)
                    .SelectMany(paragraph => paragraph.Runs)
                    .Select(run => run.Hyperlink)
                    .Any(IsUnsupportedExternalHyperlink)) {
                return $"Keynote slide {slide.Index} contains a hyperlink that cannot be represented by the PPTX owner.";
            }
            IEnumerable<IWorkGeometry> geometries = slide.TextBoxes
                .Select(textBox => textBox.Geometry)
                .Concat(slide.TitleBox == null ? Array.Empty<IWorkGeometry?>() : new[] { slide.TitleBox.Geometry })
                .Concat(slide.Images.Select(image => image.Geometry))
                .Concat(slide.Tables.Select(table => table.Geometry))
                .Where(geometry => geometry != null)
                .Cast<IWorkGeometry>();
            if (geometries.Any(geometry => Math.Abs(geometry.LeftPoints) > MaximumPointMeasurement
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
                    IWorkParagraphStyle style = paragraph.Style;
                    if (!FitsTextCoordinate(style.FirstLineIndentPoints)
                        || !FitsTextCoordinate(style.LeftIndentPoints)
                        || !FitsSpacing(style.SpaceBeforePoints)
                        || !FitsSpacing(style.SpaceAfterPoints)) {
                        return $"Keynote slide {slide.Index} contains paragraph formatting outside the PPTX range.";
                    }
                    if (paragraph.Runs.Any(run => run.Style.FontSizePoints is double fontSize
                            && (!IsFinite(fontSize) || fontSize < 1d || fontSize > 4000d))) {
                        return $"Keynote slide {slide.Index} contains a font size outside the PPTX range.";
                    }
                }
            }
            foreach (IWorkTable table in slide.Tables) {
                long tableCells = (long)table.RowCount * table.ColumnCount;
                if (table.RowCount > 4096 || table.ColumnCount > 4096
                    || tableCells > 100_000) {
                    return $"Keynote table '{table.Name}' is too large for bounded PPTX table reconstruction.";
                }
                if (destinationTableCells > MaximumDestinationTableCells - tableCells) {
                    return "Keynote tables exceed the bounded PPTX destination cell budget.";
                }
                destinationTableCells += tableCells;
                double fallbackWidth = table.DefaultColumnWidth.GetValueOrDefault(72d) * table.ColumnCount;
                double fallbackHeight = table.DefaultRowHeight.GetValueOrDefault(24d) * table.RowCount;
                if (table.Geometry == null
                    && (!FitsPositiveMeasurement(Math.Max(144d, fallbackWidth), MaximumPointMeasurement)
                        || !FitsPositiveMeasurement(Math.Max(36d, fallbackHeight), MaximumPointMeasurement))) {
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

    private static IEnumerable<IWorkTextContent> SlideText(IWorkKeynoteSlide slide) {
        if (slide.TitleBox != null) yield return slide.TitleBox.Content;
        foreach (IWorkTextBox textBox in slide.TextBoxes) yield return textBox.Content;
        yield return slide.PresenterNoteContent;
    }

    private static bool FitsTextCoordinate(double? points) => !points.HasValue
        || IsFinite(points.Value) && Math.Abs(points.Value) <= int.MaxValue / PowerPointUnits.EmusPerPoint;

    private static bool FitsSpacing(double? points) => !points.HasValue
        || IsFinite(points.Value) && points.Value >= 0 && points.Value <= int.MaxValue / 100d;

    private static bool FitsRotation(double degrees) => IsFinite(degrees)
        && Math.Abs(degrees) <= int.MaxValue / 60000d;

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool IsUnsupportedExternalHyperlink(string? value) => value != null
        && !Uri.TryCreate(value, UriKind.Absolute, out _);

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
            && Uri.TryCreate(source.Hyperlink, UriKind.Absolute, out Uri? shapeLink)) {
            textBox.SetHyperlink(shapeLink);
        }
        textBox.Clear();
        bool first = true;
        foreach (IWorkTextParagraph sourceParagraph in source.Content.Paragraphs) {
            PowerPointParagraph paragraph;
            if (first) {
                paragraph = textBox.Paragraphs[0];
                paragraph.Text = string.Empty;
                first = false;
            } else {
                paragraph = textBox.AddParagraph();
            }
            ApplyParagraphStyle(paragraph, sourceParagraph);
            WriteParagraphContent(paragraph, sourceParagraph);
        }
    }

    private static void SetRichPresenterNotes(PowerPointNotes notes, IWorkTextContent source) {
        IReadOnlyList<PowerPointParagraph> paragraphs = notes.SetParagraphs(
            source.Paragraphs.Select(_ => string.Empty));
        for (int paragraphIndex = 0; paragraphIndex < source.Paragraphs.Count; paragraphIndex++) {
            IWorkTextParagraph sourceParagraph = source.Paragraphs[paragraphIndex];
            PowerPointParagraph paragraph = paragraphs[paragraphIndex];
            ApplyParagraphStyle(paragraph, sourceParagraph);
            WriteParagraphContent(paragraph, sourceParagraph);
        }
        notes.Save();
    }

    private static void ApplyParagraphStyle(PowerPointParagraph paragraph,
        IWorkTextParagraph source) {
        IWorkParagraphStyle style = source.Style;
        if (style.Alignment.HasValue) {
            paragraph.Alignment = style.Alignment.Value switch {
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
            else if (source.ListLabel!.Length == 1) paragraph.SetBullet(source.ListLabel[0]);
        }
    }

    private static void WriteParagraphContent(PowerPointParagraph paragraph,
        IWorkTextParagraph source) {
        paragraph.Text = string.Empty;
        bool canReuseInitialRun = true;
        if (source.ListLevel >= 0 && source.ListLabel is { Length: > 1 } label) {
            AppendStyledText(paragraph, label + " ", source.Style.TextStyle, null,
                ref canReuseInitialRun);
        }
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
        if (hyperlink != null && Uri.TryCreate(hyperlink, UriKind.Absolute, out Uri? runLink)) {
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

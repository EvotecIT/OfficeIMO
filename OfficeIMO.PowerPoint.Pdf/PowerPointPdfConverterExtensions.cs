using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

/// <summary>
/// First-party PowerPoint presentation to PDF conversion helpers.
/// </summary>
public static partial class PowerPointPdfConverterExtensions {
    /// <summary>
    /// Converts a PowerPoint presentation to a first-party OfficeIMO PDF document model.
    /// </summary>
    public static PdfCore.PdfDocument ToPdfDocument(this PptCore.PowerPointPresentation presentation, PowerPointPdfSaveOptions? options = null) {
        return presentation.ToPdfDocumentResult(options).Value;
    }

    private static PdfCore.PdfDocument ConvertToPdfDocument(PptCore.PowerPointPresentation presentation, PowerPointPdfSaveOptions options) {
        if (presentation == null) {
            throw new ArgumentNullException(nameof(presentation));
        }

        PdfCore.PdfOptions pdfOptions = CreatePdfOptions(presentation, options);
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create(pdfOptions);

        if (options.PageLayout == PowerPointPdfPageLayout.NotesPages) {
            RenderNotesPages(pdf, presentation, pdfOptions.PageWidth, pdfOptions.PageHeight, options);
            return pdf;
        }
        if (options.PageLayout == PowerPointPdfPageLayout.Handouts) {
            RenderHandoutPages(pdf, presentation, pdfOptions.PageWidth, pdfOptions.PageHeight, options);
            return pdf;
        }

        IReadOnlyList<PptCore.PowerPointSlide> slides = presentation.Slides;
        int renderedSlides = 0;
        for (int slideIndex = 0; slideIndex < slides.Count; slideIndex++) {
            PptCore.PowerPointSlide slide = slides[slideIndex];
            if (!options.IncludeHiddenSlides && slide.Hidden) {
                continue;
            }

            if (renderedSlides > 0) {
                pdf.PageBreak();
            }

            RenderSlide(pdf, slide, slideIndex + 1, presentation.SlideSize.WidthPoints, presentation.SlideSize.HeightPoints, options);
            renderedSlides++;
        }

        if (renderedSlides == 0) {
            RenderEmptySlide(pdf, presentation.SlideSize.WidthPoints, presentation.SlideSize.HeightPoints);
        }

        return pdf;
    }

    /// <summary>
    /// Converts a PowerPoint presentation to a PDF document and returns conversion diagnostics with it.
    /// </summary>
    public static PdfCore.PdfDocumentConversionResult ToPdfDocumentResult(this PptCore.PowerPointPresentation presentation, PowerPointPdfSaveOptions? options = null) {
        if (presentation == null) {
            throw new ArgumentNullException(nameof(presentation));
        }

        PowerPointPdfSaveOptions operation = (options ?? new PowerPointPdfSaveOptions()).CloneForConversion();
        PdfCore.PdfDocument pdf = ConvertToPdfDocument(presentation, operation);
        return new PdfCore.PdfDocumentConversionResult(pdf, operation.Report);
    }

    /// <summary>
    /// Converts a PowerPoint presentation to PDF bytes.
    /// </summary>
    /// <example><code>byte[] pdf = presentation.ToPdf();</code></example>
    public static byte[] ToPdf(this PptCore.PowerPointPresentation presentation, PowerPointPdfSaveOptions? options = null) {
        return presentation.ToPdfDocument(options).ToBytes();
    }

    /// <summary>
    /// Saves a PowerPoint presentation as a PDF file.
    /// </summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this PptCore.PowerPointPresentation presentation, string path, PowerPointPdfSaveOptions? options = null) =>
        presentation.ToPdfDocumentResult(options).Save(path);

    /// <summary>
    /// Attempts to save a PowerPoint presentation as a PDF file and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this PptCore.PowerPointPresentation presentation, string path, PowerPointPdfSaveOptions? options = null) {
        try {
            return presentation.ToPdfDocumentResult(options).TrySave(path);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>
    /// Writes a PowerPoint presentation as PDF to a stream.
    /// </summary>
    public static PdfCore.PdfSaveResult SaveAsPdf(this PptCore.PowerPointPresentation presentation, Stream stream, PowerPointPdfSaveOptions? options = null) =>
        presentation.ToPdfDocumentResult(options).Save(stream);

    /// <summary>
    /// Attempts to write a PowerPoint presentation as PDF to a stream and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this PptCore.PowerPointPresentation presentation, Stream stream, PowerPointPdfSaveOptions? options = null) {
        try {
            return presentation.ToPdfDocumentResult(options).TrySave(stream);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    /// <summary>Converts synchronously, then asynchronously saves a PowerPoint presentation PDF at the specified path.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this PptCore.PowerPointPresentation presentation,
        string path,
        PowerPointPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return presentation.ToPdfDocumentResult(options).SaveAsync(path, cancellationToken);
    }

    /// <summary>Converts synchronously, then asynchronously saves a PowerPoint presentation PDF to a caller-owned stream.</summary>
    public static Task<PdfCore.PdfSaveResult> SaveAsPdfAsync(
        this PptCore.PowerPointPresentation presentation,
        Stream stream,
        PowerPointPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        return presentation.ToPdfDocumentResult(options).SaveAsync(stream, cancellationToken);
    }

    /// <summary>Attempts to asynchronously save a PowerPoint presentation as PDF at the specified path.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(
        this PptCore.PowerPointPresentation presentation,
        string path,
        PowerPointPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await presentation.ToPdfDocumentResult(options)
                .TrySaveAsync(path, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>Attempts to asynchronously save a PowerPoint presentation as PDF to a caller-owned stream.</summary>
    public static async Task<PdfCore.PdfSaveResult> TrySaveAsPdfAsync(
        this PptCore.PowerPointPresentation presentation,
        Stream stream,
        PowerPointPdfSaveOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return await presentation.ToPdfDocumentResult(options)
                .TrySaveAsync(stream, cancellationToken)
                .ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    private static void RenderSlide(PdfCore.PdfDocument pdf, PptCore.PowerPointSlide slide, int slideNumber, double pageWidth, double pageHeight, PowerPointPdfSaveOptions options) {
        pdf.Canvas(canvas => {
            RenderSlideBackground(canvas, slide, slideNumber, pageWidth, pageHeight, options);

            RenderShapes(canvas, slide.GetInheritedShapesForExport(), slideNumber, pageWidth, pageHeight, options, warnInvalidBounds: false, groupDepth: 0);
            RenderShapes(canvas, slide.Shapes, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds: true, groupDepth: 0);
        });
    }

    private static void RenderNotesPages(PdfCore.PdfDocument pdf, PptCore.PowerPointPresentation presentation,
        double pageWidth, double pageHeight, PowerPointPdfSaveOptions options) {
        List<(PptCore.PowerPointSlide Slide, int Number)> slides = GetVisibleSlides(presentation, options);
        if (slides.Count == 0) {
            RenderEmptySlide(pdf, pageWidth, pageHeight);
            return;
        }
        for (int index = 0; index < slides.Count; index++) {
            if (index > 0) pdf.PageBreak();
            var item = slides[index];
            pdf.Canvas(canvas => {
                double margin = 54D;
                double availableWidth = pageWidth - margin * 2D;
                double slideHeight = availableWidth * presentation.SlideSize.HeightPoints /
                    presentation.SlideSize.WidthPoints;
                RenderSnapshotThumbnail(canvas, item.Slide, item.Number, margin, margin, availableWidth,
                    slideHeight, options);
                canvas.Text("Slide " + item.Number.ToString(CultureInfo.InvariantCulture), margin,
                    margin + slideHeight + 18D, availableWidth, 18D, fontSize: 10D,
                    color: PdfCore.PdfColor.FromRgb(80, 86, 96));
                string notes = string.Empty;
                if (options.IncludeSpeakerNotes) item.Slide.Notes.TryGetText(out notes);
                canvas.Text(string.IsNullOrWhiteSpace(notes) ? "No speaker notes." : notes,
                    margin, margin + slideHeight + 48D, availableWidth,
                    Math.Max(36D, pageHeight - margin * 2D - slideHeight - 48D), fontSize: 12D,
                    color: PdfCore.PdfColor.FromRgb(35, 42, 52));
            });
        }
    }

    private static void RenderHandoutPages(PdfCore.PdfDocument pdf, PptCore.PowerPointPresentation presentation,
        double pageWidth, double pageHeight, PowerPointPdfSaveOptions options) {
        List<(PptCore.PowerPointSlide Slide, int Number)> slides = GetVisibleSlides(presentation, options);
        if (slides.Count == 0) {
            RenderEmptySlide(pdf, pageWidth, pageHeight);
            return;
        }
        int perPage = options.HandoutSlidesPerPage;
        for (int pageStart = 0, pageIndex = 0; pageStart < slides.Count; pageStart += perPage, pageIndex++) {
            if (pageIndex > 0) pdf.PageBreak();
            List<(PptCore.PowerPointSlide Slide, int Number)> pageSlides = slides.Skip(pageStart)
                .Take(perPage).ToList();
            pdf.Canvas(canvas => RenderHandoutPage(canvas, presentation, pageSlides, pageWidth, pageHeight, options));
        }
    }

    private static void RenderHandoutPage(PdfCore.PdfPageCanvas canvas,
        PptCore.PowerPointPresentation presentation,
        IReadOnlyList<(PptCore.PowerPointSlide Slide, int Number)> slides,
        double pageWidth, double pageHeight, PowerPointPdfSaveOptions options) {
        const double margin = 36D;
        const double gutter = 18D;
        if (options.HandoutSlidesPerPage == 3) {
            double rowHeight = (pageHeight - margin * 2D - gutter * 2D) / 3D;
            double thumbWidth = (pageWidth - margin * 2D - gutter) * 0.46D;
            for (int index = 0; index < slides.Count; index++) {
                double top = margin + index * (rowHeight + gutter);
                double thumbHeight = Math.Min(rowHeight - 18D, thumbWidth *
                    presentation.SlideSize.HeightPoints / presentation.SlideSize.WidthPoints);
                RenderSnapshotThumbnail(canvas, slides[index].Slide, slides[index].Number,
                    margin, top, thumbWidth, thumbHeight, options);
                string notes = string.Empty;
                if (options.IncludeSpeakerNotes) slides[index].Slide.Notes.TryGetText(out notes);
                string noteText = string.IsNullOrWhiteSpace(notes)
                    ? "Slide " + slides[index].Number.ToString(CultureInfo.InvariantCulture) +
                      Environment.NewLine + "________________________________" + Environment.NewLine +
                      "________________________________" + Environment.NewLine + "________________________________"
                    : "Slide " + slides[index].Number.ToString(CultureInfo.InvariantCulture) +
                      Environment.NewLine + notes;
                canvas.Text(noteText, margin + thumbWidth + gutter, top,
                    pageWidth - margin * 2D - thumbWidth - gutter, rowHeight,
                    fontSize: 9D, color: PdfCore.PdfColor.FromRgb(55, 62, 72));
            }
            return;
        }

        (int columns, int rows) = GetHandoutGrid(options.HandoutSlidesPerPage);
        double cellWidth = (pageWidth - margin * 2D - gutter * (columns - 1)) / columns;
        double cellHeight = (pageHeight - margin * 2D - gutter * (rows - 1)) / rows;
        for (int index = 0; index < slides.Count; index++) {
            int row = index / columns;
            int column = index % columns;
            double left = margin + column * (cellWidth + gutter);
            double top = margin + row * (cellHeight + gutter);
            double notesHeight = options.IncludeSpeakerNotes ? Math.Min(36D, cellHeight * 0.22D) : 0D;
            double labelHeight = 14D;
            double thumbnailAreaHeight = cellHeight - labelHeight - notesHeight;
            double thumbHeight = Math.Min(thumbnailAreaHeight, cellWidth *
                presentation.SlideSize.HeightPoints / presentation.SlideSize.WidthPoints);
            double thumbWidth = thumbHeight * presentation.SlideSize.WidthPoints /
                presentation.SlideSize.HeightPoints;
            RenderSnapshotThumbnail(canvas, slides[index].Slide, slides[index].Number,
                left + (cellWidth - thumbWidth) / 2D, top, thumbWidth, thumbHeight, options);
            canvas.Text("Slide " + slides[index].Number.ToString(CultureInfo.InvariantCulture),
                left, top + thumbHeight + 2D, cellWidth, labelHeight, fontSize: 8D,
                color: PdfCore.PdfColor.FromRgb(80, 86, 96), align: PdfCore.PdfAlign.Center);
            if (notesHeight > 0D && slides[index].Slide.Notes.TryGetText(out string notes) &&
                !string.IsNullOrWhiteSpace(notes)) {
                canvas.Text(notes, left, top + thumbHeight + labelHeight + 2D, cellWidth, notesHeight,
                    fontSize: 7D, color: PdfCore.PdfColor.FromRgb(55, 62, 72));
            }
        }
    }

    private static void RenderSnapshotThumbnail(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointSlide slide,
        int slideNumber, double x, double y, double width, double height, PowerPointPdfSaveOptions options) {
        PptCore.PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
            new PptCore.PowerPointImageExportOptions {
                IncludeSlideBackground = options.IncludeSlideBackgrounds,
                IncludePictures = options.IncludePictures,
                IncludeAutoShapes = options.IncludeAutoShapes,
                IncludeTextBoxes = options.IncludeTextBoxes,
                IncludeTables = options.IncludeTables,
                IncludeCharts = options.IncludeCharts
            });
        canvas.Drawing(snapshot.Drawing, x, y, width, height);
        foreach (OfficeImageExportDiagnostic diagnostic in snapshot.Diagnostics) {
            AddWarning(options, slideNumber, "snapshot-" + diagnostic.Code, diagnostic.Message);
        }
    }

    private static List<(PptCore.PowerPointSlide Slide, int Number)> GetVisibleSlides(
        PptCore.PowerPointPresentation presentation, PowerPointPdfSaveOptions options) {
        var slides = new List<(PptCore.PowerPointSlide, int)>();
        for (int index = 0; index < presentation.Slides.Count; index++) {
            PptCore.PowerPointSlide slide = presentation.Slides[index];
            if (!options.IncludeHiddenSlides && slide.Hidden) continue;
            slides.Add((slide, index + 1));
        }
        return slides;
    }

    private static (int Columns, int Rows) GetHandoutGrid(int slidesPerPage) {
        switch (slidesPerPage) {
            case 1: return (1, 1);
            case 2: return (1, 2);
            case 4: return (2, 2);
            case 6: return (2, 3);
            case 9: return (3, 3);
            default: return (1, 3);
        }
    }

    private static void RenderShapes(PdfCore.PdfPageCanvas canvas, IReadOnlyList<PptCore.PowerPointShape> shapes, int slideNumber, double pageWidth, double pageHeight, PowerPointPdfSaveOptions options, bool warnInvalidBounds, int groupDepth) {
        foreach (PptCore.PowerPointShape shape in shapes) {
            if (shape.Hidden) {
                continue;
            }

            if (!TryGetShapeBox(shape, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, out double x, out double y, out double width, out double height)) {
                continue;
            }

            Action<PdfCore.PdfPageCanvas> render = target => RenderShapeContent(target, shape, x, y, width, height, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, groupDepth);
            if (TryGetVisibleSlideBox(x, y, width, height, pageWidth, pageHeight, out double clipX, out double clipY, out double clipWidth, out double clipHeight) &&
                NeedsSlideClip(x, y, width, height, pageWidth, pageHeight)) {
                canvas.Clip(clipX, clipY, clipWidth, clipHeight, render);
            } else {
                render(canvas);
            }
        }
    }

    private static void RenderShapeContent(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointShape shape, double x, double y, double width, double height, int slideNumber, double pageWidth, double pageHeight, PowerPointPdfSaveOptions options, bool warnInvalidBounds, int groupDepth) {
        if (shape is PptCore.PowerPointTextBox textBox) {
            bool renderedGeometry = options.IncludeAutoShapes && RenderTextBoxGeometry(canvas, textBox, x, y, width, height);
            if (options.IncludeTextBoxes) {
                RenderTextBox(canvas, textBox, x, y, width, height, slideNumber, options, suppressFrame: renderedGeometry);
            }
            return;
        }

        if (shape is PptCore.PowerPointPicture picture) {
            if (options.IncludePictures) {
                RenderPicture(canvas, picture, x, y, width, height, slideNumber, options);
            }
            return;
        }

        if (shape is PptCore.PowerPointTable table) {
            if (options.IncludeTables) {
                RenderTable(canvas, table, x, y, width, height, slideNumber, options);
            }
            return;
        }

        if (shape is PptCore.PowerPointChart chart) {
            if (options.IncludeCharts) {
                RenderChart(canvas, chart, x, y, width, height, slideNumber, options);
            }
            return;
        }

        if (shape is PptCore.PowerPointAutoShape autoShape) {
            if (options.IncludeAutoShapes) {
                RenderAutoShape(canvas, autoShape, x, y, width, height, slideNumber, options);
            }
            return;
        }

        if (shape is PptCore.PowerPointGroupShape groupShape) {
            if (shape.OwnerSlide != null) {
                RenderGroupShape(canvas, groupShape, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, groupDepth);
            } else {
                AddWarning(options, slideNumber, "unsupported-shape", "Skipped a PowerPoint group shape because its owning slide context could not be resolved.");
            }
            return;
        }

        AddWarning(options, slideNumber, "unsupported-shape", "Skipped unsupported PowerPoint shape content type '" + shape.ShapeContentType + "'.");
    }

    private static void RenderGroupShape(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointGroupShape groupShape, int slideNumber, double pageWidth, double pageHeight, PowerPointPdfSaveOptions options, bool warnInvalidBounds, int groupDepth) {
        if (options.MaxGroupShapeDepth >= 0 && groupDepth >= options.MaxGroupShapeDepth) {
            AddWarning(options, slideNumber, "group-depth-limit", "Skipped nested PowerPoint group shape content because MaxGroupShapeDepth was reached.");
            return;
        }

        IReadOnlyList<PptCore.PowerPointShape> children = groupShape.OwnerSlide!.GetGroupChildren(groupShape);
        foreach (PptCore.PowerPointShape child in children) {
            if (child.Hidden) {
                continue;
            }

            if (!TryGetShapeBox(child, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, out double x, out double y, out double width, out double height)) {
                continue;
            }

            MapGroupChildBox(groupShape, ref x, ref y, ref width, ref height);
            Action<PdfCore.PdfPageCanvas> render = target => RenderShapeContent(target, child, x, y, width, height, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, groupDepth + 1);
            if (TryGetVisibleSlideBox(x, y, width, height, pageWidth, pageHeight, out double clipX, out double clipY, out double clipWidth, out double clipHeight) &&
                NeedsSlideClip(x, y, width, height, pageWidth, pageHeight)) {
                canvas.Clip(clipX, clipY, clipWidth, clipHeight, render);
            } else {
                render(canvas);
            }
        }
    }

    private static void MapGroupChildBox(PptCore.PowerPointGroupShape groupShape, ref double x, ref double y, ref double width, ref double height) {
        TransformGroup? transform = groupShape.GroupShape.GroupShapeProperties?.TransformGroup;
        long? groupXEmu = transform?.Offset?.X?.Value;
        long? groupYEmu = transform?.Offset?.Y?.Value;
        long? groupWidthEmu = transform?.Extents?.Cx?.Value;
        long? groupHeightEmu = transform?.Extents?.Cy?.Value;
        long? childXEmu = transform?.ChildOffset?.X?.Value;
        long? childYEmu = transform?.ChildOffset?.Y?.Value;
        long? childWidthEmu = transform?.ChildExtents?.Cx?.Value;
        long? childHeightEmu = transform?.ChildExtents?.Cy?.Value;
        if (!groupXEmu.HasValue || !groupYEmu.HasValue || !groupWidthEmu.HasValue || !groupHeightEmu.HasValue ||
            !childXEmu.HasValue || !childYEmu.HasValue || !childWidthEmu.HasValue || !childHeightEmu.HasValue ||
            childWidthEmu.Value == 0L || childHeightEmu.Value == 0L) {
            return;
        }

        double groupX = PptCore.PowerPointUnits.ToPoints(groupXEmu.Value);
        double groupY = PptCore.PowerPointUnits.ToPoints(groupYEmu.Value);
        double childX = PptCore.PowerPointUnits.ToPoints(childXEmu.Value);
        double childY = PptCore.PowerPointUnits.ToPoints(childYEmu.Value);
        double scaleX = groupWidthEmu.Value / (double)childWidthEmu.Value;
        double scaleY = groupHeightEmu.Value / (double)childHeightEmu.Value;
        x = groupX + (x - childX) * scaleX;
        y = groupY + (y - childY) * scaleY;
        width *= scaleX;
        height *= scaleY;
    }

    private static void RenderEmptySlide(PdfCore.PdfDocument pdf, double pageWidth, double pageHeight) {
        pdf.Canvas(canvas => RenderFallbackBackground(canvas, pageWidth, pageHeight));
    }

    private static void RenderSlideBackground(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointSlide slide, int slideNumber, double pageWidth, double pageHeight, PowerPointPdfSaveOptions options) {
        if (!options.IncludeSlideBackgrounds) {
            RenderFallbackBackground(canvas, pageWidth, pageHeight);
            return;
        }

        PptCore.PowerPointSlideBackground background = slide.GetBackground();
        if (background.Kind == PptCore.PowerPointSlideBackgroundKind.SolidColor && !string.IsNullOrWhiteSpace(background.Color)) {
            OfficeColor? color = ParseOfficeColor(background.Color);
            if (color.HasValue) {
                RenderBackgroundShape(canvas, pageWidth, pageHeight, color.Value, fillGradient: null);
                return;
            }

            AddWarning(options, slideNumber, "invalid-slide-background-color", "Used a white fallback because the PowerPoint slide background color could not be parsed.");
        } else if (background.Kind == PptCore.PowerPointSlideBackgroundKind.Image && background.ImageBytes != null) {
            RenderFallbackBackground(canvas, pageWidth, pageHeight);
            try {
                var imageStyle = new PdfCore.PdfImageStyle { Fit = OfficeImageFit.Stretch };
                if (background.HasImageCrop) {
                    imageStyle.SourceCrop = new PdfCore.PdfImageSourceCrop(background.ImageCropLeft, background.ImageCropTop, background.ImageCropRight, background.ImageCropBottom);
                }

                canvas.Image(
                    background.ImageBytes,
                    0,
                    0,
                    pageWidth,
                    pageHeight,
                    style: imageStyle,
                    alternativeText: "PowerPoint slide background");
                return;
            } catch (Exception ex) {
                AddWarning(options, slideNumber, "unsupported-slide-background-image", "Used a white fallback because the PowerPoint slide background image could not be embedded as a PDF image: " + ex.Message);
                return;
            }
        } else if (background.Kind == PptCore.PowerPointSlideBackgroundKind.LinearGradient) {
            string gradientStartColor = background.GradientStartColor ?? string.Empty;
            string gradientEndColor = background.GradientEndColor ?? string.Empty;
            OfficeLinearGradient? gradient = CreateLinearGradient(gradientStartColor, gradientEndColor, background.GradientAngleDegrees ?? 0D);
            if (gradient != null) {
                OfficeColor fallbackColor = gradient.Stops[0].Color;
                RenderBackgroundShape(canvas, pageWidth, pageHeight, fallbackColor, gradient);
                return;
            }

            AddWarning(options, slideNumber, "invalid-slide-background-gradient", "Used a white fallback because the PowerPoint slide background gradient colors could not be parsed.");
        } else if (background.Kind == PptCore.PowerPointSlideBackgroundKind.Unsupported) {
            AddWarning(options, slideNumber, "unsupported-slide-background", "Used a white fallback because " + (background.UnsupportedReason ?? "the PowerPoint slide background is not supported."));
        }

        RenderFallbackBackground(canvas, pageWidth, pageHeight);
    }

    private static void RenderFallbackBackground(PdfCore.PdfPageCanvas canvas, double pageWidth, double pageHeight) {
        RenderBackgroundShape(canvas, pageWidth, pageHeight, OfficeColor.White, fillGradient: null);
    }

    private static void RenderBackgroundShape(PdfCore.PdfPageCanvas canvas, double pageWidth, double pageHeight, OfficeColor fillColor, OfficeLinearGradient? fillGradient) {
        OfficeShape background = OfficeShape.Rectangle(pageWidth, pageHeight);
        background.FillColor = fillColor;
        background.FillGradient = fillGradient?.Clone();
        background.StrokeWidth = 0D;
        canvas.Shape(background, 0, 0);
    }

    private static void RenderTextBox(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointTextBox textBox, double x, double y, double width, double height, int slideNumber, PowerPointPdfSaveOptions options, bool suppressFrame = false) {
        PdfCore.PdfCanvasTextBoxStyle style = CreateTextBoxStyle(textBox);
        if (!TryFitTextBoxPadding(style, width, height, out bool adjustedPadding)) {
            AddWarning(options, slideNumber, "invalid-text-box-bounds", "Skipped a PowerPoint text box because its margins leave no renderable PDF text area.");
            return;
        }

        if (adjustedPadding) {
            AddWarning(options, slideNumber, "text-box-padding", "Reduced PowerPoint text box margins because the original margins left no renderable PDF text area.");
        }

        if (suppressFrame) {
            RemoveTextBoxFrame(style);
        }

        if (ShouldRenderParagraphsIndividually(textBox) && (textBox.Rotation ?? 0D) == 0D) {
            RenderParagraphTextBox(canvas, textBox, x, y, width, height, style, slideNumber, options, renderFrame: !suppressFrame);
            return;
        }

        IReadOnlyList<PdfCore.TextRun> runs = CreateTextRuns(textBox, slideNumber, options);
        bool reportedOverflow = false;
        canvas.TextBox(runs, x, y, width, height, style, textBox.Rotation ?? 0D, diagnostic => {
            if (reportedOverflow || diagnostic.Kind != PdfCore.PdfLayoutDiagnosticKind.ClippedContent) {
                return;
            }

            reportedOverflow = true;
            AddWarning(
                options,
                slideNumber,
                "text-box-overflow",
                "Clipped PowerPoint text box content because the PDF text box render pass found more text than fits the mapped text area.",
                new PdfCore.PdfLayoutDiagnostic(
                    PdfCore.PdfLayoutDiagnosticKind.ClippedContent,
                    "PowerPointTextBox",
                    diagnostic.Message,
                    x,
                    y,
                    width,
                    height));
        });
    }

    private static bool ShouldRenderParagraphsIndividually(PptCore.PowerPointTextBox textBox) {
        IReadOnlyList<PptCore.PowerPointParagraph> paragraphs = textBox.Paragraphs;
        if (paragraphs.Count <= 1) {
            return HasListMarker(paragraphs.FirstOrDefault());
        }

        TextAlignmentTypeValues? firstAlignment = paragraphs[0].Alignment;
        return paragraphs.Any(paragraph => paragraph.Alignment != firstAlignment || HasListMarker(paragraph));
    }

    private static bool HasListMarker(PptCore.PowerPointParagraph? paragraph) =>
        paragraph != null && (!string.IsNullOrEmpty(paragraph.BulletCharacter) || paragraph.IsNumbered);

    private static void RenderParagraphTextBox(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointTextBox textBox, double x, double y, double width, double height, PdfCore.PdfCanvasTextBoxStyle style, int slideNumber, PowerPointPdfSaveOptions options, bool renderFrame) {
        if (renderFrame) {
            RenderTextBoxFrame(canvas, textBox, x, y, width, height);
        }

        double paddingLeft = style.PaddingLeft ?? style.PaddingX;
        double paddingRight = style.PaddingRight ?? style.PaddingX;
        double paddingTop = style.PaddingTop ?? style.PaddingY;
        double paddingBottom = style.PaddingBottom ?? style.PaddingY;
        double textX = x + paddingLeft;
        double textY = y + paddingTop;
        double textWidth = Math.Max(1D, width - paddingLeft - paddingRight);
        double textHeight = Math.Max(1D, height - paddingTop - paddingBottom);
        double fontSize = style.FontSize ?? 12D;
        double lineHeight = style.LineHeight ?? fontSize * 1.2D;
        AddPowerPointListLayoutDiagnostics(options, slideNumber, textBox, x, y, width, height);
        var paragraphRuns = new List<IReadOnlyList<PdfCore.TextRun>>();
        var paragraphHeights = new List<double>();
        int numberIndex = 1;

        foreach (PptCore.PowerPointParagraph paragraph in textBox.Paragraphs) {
            IReadOnlyList<PdfCore.TextRun> runs = CreateParagraphRuns(paragraph, textBox, slideNumber, options, ref numberIndex);
            paragraphRuns.Add(runs);
            paragraphHeights.Add(EstimateParagraphHeight(runs, textWidth, fontSize, lineHeight));
        }

        if (paragraphRuns.Count == 0) {
            paragraphRuns.Add(new[] { PdfCore.TextRun.Normal(string.Empty) });
            paragraphHeights.Add(lineHeight);
        }

        double totalHeight = paragraphHeights.Sum();
        double offsetY = style.VerticalAlign switch {
            PdfCore.PdfVerticalAlign.Middle => Math.Max(0D, textHeight - totalHeight) / 2D,
            PdfCore.PdfVerticalAlign.Bottom => Math.Max(0D, textHeight - totalHeight),
            _ => 0D
        };

        double cursorY = textY + offsetY;
        int renderedParagraphs = 0;
        bool reportedOverflow = false;
        Action<PdfCore.PdfLayoutDiagnostic?> reportOverflow = diagnostic => {
            if (reportedOverflow) {
                return;
            }

            reportedOverflow = true;
            AddWarning(
                options,
                slideNumber,
                "text-box-overflow",
                "Clipped PowerPoint text box content because the PDF text box render pass found more text than fits the mapped text area.",
                new PdfCore.PdfLayoutDiagnostic(
                    PdfCore.PdfLayoutDiagnosticKind.ClippedContent,
                    "PowerPointTextBox",
                    diagnostic?.Message ?? "The mapped PDF text area was too small for all PowerPoint paragraphs.",
                    x,
                    y,
                    width,
                    height));
        };

        for (int index = 0; index < paragraphRuns.Count && cursorY < textY + textHeight; index++) {
            PptCore.PowerPointParagraph paragraph = textBox.Paragraphs.Count > index ? textBox.Paragraphs[index] : textBox.Paragraphs.Last();
            double availableHeight = textY + textHeight - cursorY;
            double paragraphHeight = Math.Min(paragraphHeights[index], availableHeight);
            var paragraphStyle = CreateTransparentParagraphStyle(style, paragraph);
            canvas.TextBox(paragraphRuns[index], textX, cursorY, textWidth, Math.Max(1D, paragraphHeight), paragraphStyle, diagnosticHandler: diagnostic => {
                if (diagnostic.Kind == PdfCore.PdfLayoutDiagnosticKind.ClippedContent) {
                    reportOverflow(diagnostic);
                }
            });
            cursorY += paragraphHeight;
            renderedParagraphs++;
        }

        if (renderedParagraphs < paragraphRuns.Count) {
            reportOverflow(null);
        }
    }

    private static void RemoveTextBoxFrame(PdfCore.PdfCanvasTextBoxStyle style) {
        style.Background = null;
        style.BackgroundOpacity = null;
        style.BorderColor = null;
        style.BorderWidth = 0D;
        style.BorderDashStyle = OfficeStrokeDashStyle.Solid;
    }

    private static void RenderTextBoxFrame(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointTextBox textBox, double x, double y, double width, double height) {
        OfficeColor? fill = textBox.FillTransparency == 100 ? null : ParseOfficeColor(textBox.FillColor);
        OfficeColor? outline = ParseOfficeColor(textBox.OutlineColor);
        if (!fill.HasValue && !outline.HasValue) {
            return;
        }

        OfficeShape frame = OfficeShape.Rectangle(width, height);
        frame.FillColor = fill;
        if (textBox.FillTransparency.HasValue && textBox.FillTransparency.Value > 0 && textBox.FillTransparency.Value < 100) {
            frame.FillOpacity = 1D - textBox.FillTransparency.Value / 100D;
        }

        frame.StrokeColor = outline;
        frame.StrokeWidth = outline.HasValue ? textBox.OutlineWidthPoints ?? 0.75D : 0D;
        frame.StrokeDashStyle = MapDash(textBox.OutlineDash);
        canvas.Shape(frame, x, y, rotationAngle: textBox.Rotation ?? 0D);
    }

    private static PdfCore.PdfCanvasTextBoxStyle CreateTransparentParagraphStyle(PdfCore.PdfCanvasTextBoxStyle style, PptCore.PowerPointParagraph paragraph) {
        PdfCore.PdfCanvasTextBoxStyle paragraphStyle = style.Clone();
        paragraphStyle.Background = null;
        paragraphStyle.BorderColor = null;
        paragraphStyle.BorderWidth = 0D;
        paragraphStyle.PaddingX = 0D;
        paragraphStyle.PaddingY = 0D;
        paragraphStyle.PaddingLeft = 0D;
        paragraphStyle.PaddingRight = 0D;
        paragraphStyle.PaddingTop = 0D;
        paragraphStyle.PaddingBottom = 0D;
        paragraphStyle.Align = MapAlign(paragraph.Alignment);
        paragraphStyle.VerticalAlign = PdfCore.PdfVerticalAlign.Top;
        return paragraphStyle;
    }

    private static double EstimateParagraphHeight(IReadOnlyList<PdfCore.TextRun> runs, double width, double fontSize, double lineHeight) {
        int lineCount = 0;
        double averageCharacterWidth = Math.Max(1D, fontSize * 0.52D);
        int charactersPerLine = Math.Max(1, (int)Math.Floor(width / averageCharacterWidth));
        string text = string.Concat(runs.Select(run => run.Text));
        foreach (string line in text.Split('\n')) {
            lineCount += Math.Max(1, (int)Math.Ceiling(line.Length / (double)charactersPerLine));
        }

        return Math.Max(lineHeight, lineCount * lineHeight);
    }

    private static bool TryFitTextBoxPadding(PdfCore.PdfCanvasTextBoxStyle style, double width, double height, out bool adjustedPadding) {
        const double minimumInnerSize = 0.5D;
        double originalPaddingLeft = style.PaddingLeft ?? style.PaddingX;
        double originalPaddingRight = style.PaddingRight ?? style.PaddingX;
        double originalPaddingTop = style.PaddingTop ?? style.PaddingY;
        double originalPaddingBottom = style.PaddingBottom ?? style.PaddingY;
        adjustedPadding = false;
        if (width <= minimumInnerSize || height <= minimumInnerSize) {
            return false;
        }

        FitPaddingPair(width, minimumInnerSize, originalPaddingLeft, originalPaddingRight, out double paddingLeft, out double paddingRight);
        FitPaddingPair(height, minimumInnerSize, originalPaddingTop, originalPaddingBottom, out double paddingTop, out double paddingBottom);
        style.PaddingLeft = paddingLeft;
        style.PaddingRight = paddingRight;
        style.PaddingTop = paddingTop;
        style.PaddingBottom = paddingBottom;
        adjustedPadding =
            paddingLeft != originalPaddingLeft ||
            paddingRight != originalPaddingRight ||
            paddingTop != originalPaddingTop ||
            paddingBottom != originalPaddingBottom;
        return paddingLeft + paddingRight < width && paddingTop + paddingBottom < height;
    }

    private static void FitPaddingPair(double outerSize, double minimumInnerSize, double leading, double trailing, out double fittedLeading, out double fittedTrailing) {
        double maximumPadding = Math.Max(0D, outerSize - minimumInnerSize);
        double total = leading + trailing;
        if (total <= maximumPadding || total <= 0D) {
            fittedLeading = leading;
            fittedTrailing = trailing;
            return;
        }

        double scale = maximumPadding / total;
        fittedLeading = leading * scale;
        fittedTrailing = trailing * scale;
    }

    private static IReadOnlyList<PdfCore.TextRun> CreateTextRuns(PptCore.PowerPointTextBox textBox, int slideNumber, PowerPointPdfSaveOptions options) {
        var runs = new List<PdfCore.TextRun>();
        IReadOnlyList<PptCore.PowerPointParagraph> paragraphs = textBox.Paragraphs;
        int numberIndex = 1;
        for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++) {
            if (paragraphIndex > 0) {
                runs.Add(PdfCore.TextRun.LineBreak());
            }

            runs.AddRange(CreateParagraphRuns(paragraphs[paragraphIndex], textBox, slideNumber, options, ref numberIndex));
        }

        if (runs.Count == 0) {
            runs.Add(PdfCore.TextRun.Normal(string.Empty));
        }

        return runs;
    }

    private static IReadOnlyList<PdfCore.TextRun> CreateParagraphRuns(PptCore.PowerPointParagraph paragraph, PptCore.PowerPointTextBox textBox, int slideNumber, PowerPointPdfSaveOptions options, ref int numberIndex) {
        var runs = new List<PdfCore.TextRun>();
        string prefix = CreateListPrefix(paragraph, ref numberIndex);
        if (!string.IsNullOrEmpty(prefix)) {
            runs.Add(PdfCore.TextRun.Normal(prefix, ParsePdfColor(textBox.Color), textBox.FontSize, font: MapFont(textBox.FontName)));
        }

        IReadOnlyList<PptCore.PowerPointTextRun> paragraphRuns = paragraph.Runs;
        if (paragraphRuns.Count == 0 && !paragraph.Paragraph.ChildElements.Any(child => child is A.Break or A.Field)) {
            runs.Add(new PdfCore.TextRun(paragraph.Text));
            return runs;
        }

        int runIndex = 0;
        bool hasInlineContent = false;
        foreach (OpenXmlElement child in paragraph.Paragraph.ChildElements) {
            switch (child) {
                case A.Run:
                    if (runIndex < paragraphRuns.Count) {
                        runs.Add(CreateTextRun(paragraphRuns[runIndex], textBox, slideNumber, options));
                        hasInlineContent = true;
                    }

                    runIndex++;
                    break;
                case A.Break:
                    runs.Add(PdfCore.TextRun.LineBreak());
                    hasInlineContent = true;
                    break;
                case A.Field field:
                    string fieldText = field.Text?.Text ?? field.InnerText ?? string.Empty;
                    if (!string.IsNullOrEmpty(fieldText)) {
                        runs.Add(PdfCore.TextRun.Normal(fieldText, ParsePdfColor(textBox.Color), textBox.FontSize, font: MapFont(textBox.FontName)));
                        hasInlineContent = true;
                    }

                    break;
            }
        }

        if (!hasInlineContent) {
            runs.Add(new PdfCore.TextRun(paragraph.Text));
        }

        return runs;
    }

    private static string CreateListPrefix(PptCore.PowerPointParagraph paragraph, ref int numberIndex) {
        string indent = paragraph.Level.HasValue && paragraph.Level.Value > 0
            ? new string(' ', Math.Min(18, paragraph.Level.Value * 2))
            : string.Empty;
        if (!string.IsNullOrEmpty(paragraph.BulletCharacter)) {
            return indent + paragraph.BulletCharacter + " ";
        }

        if (paragraph.IsNumbered) {
            int number = paragraph.NumberingStartAt ?? numberIndex;
            numberIndex = number + 1;
            return indent + number.ToString(System.Globalization.CultureInfo.InvariantCulture) + ". ";
        }

        return string.Empty;
    }

    private static PdfCore.TextRun CreateTextRun(PptCore.PowerPointTextRun run, PptCore.PowerPointTextBox textBox, int slideNumber, PowerPointPdfSaveOptions options) {
        string text = run.Text ?? string.Empty;
        PdfCore.PdfColor? color = ParsePdfColor(run.Color ?? textBox.Color);
        PdfCore.PdfStandardFont? font = MapFont(run.FontName ?? textBox.FontName);
        double? fontSize = run.FontSize ?? textBox.FontSize;
        Uri? hyperlink = run.Hyperlink;
        string? linkUri = hyperlink != null && hyperlink.IsAbsoluteUri && !string.IsNullOrEmpty(text) ? hyperlink.AbsoluteUri : null;
        if (hyperlink != null && !hyperlink.IsAbsoluteUri) {
            AddWarning(options, slideNumber, "relative-hyperlink", "Skipped a relative PowerPoint hyperlink because PDF URI annotations require absolute targets.");
        }

        return new PdfCore.TextRun(
            text,
            bold: run.Bold,
            underline: run.Underline || linkUri != null,
            color: color,
            italic: run.Italic,
            fontSize: fontSize,
            font: font,
            linkUri: linkUri);
    }

    private static PdfCore.PdfCanvasTextBoxStyle CreateTextBoxStyle(PptCore.PowerPointTextBox textBox) {
        PdfCore.PdfColor? fill = ParsePdfColor(textBox.FillColor);
        PdfCore.PdfColor? outline = ParsePdfColor(textBox.OutlineColor);
        return new PdfCore.PdfCanvasTextBoxStyle {
            Background = textBox.FillTransparency == 100 ? null : fill,
            BackgroundOpacity = textBox.FillTransparency.HasValue && textBox.FillTransparency.Value > 0 && textBox.FillTransparency.Value < 100
                ? 1D - textBox.FillTransparency.Value / 100D
                : null,
            BorderColor = outline,
            BorderWidth = outline.HasValue ? textBox.OutlineWidthPoints ?? 0.75D : 0D,
            BorderDashStyle = MapDash(textBox.OutlineDash),
            PaddingLeft = textBox.TextMarginLeftPoints ?? 3.6D,
            PaddingRight = textBox.TextMarginRightPoints ?? 3.6D,
            PaddingTop = textBox.TextMarginTopPoints ?? 3.6D,
            PaddingBottom = textBox.TextMarginBottomPoints ?? 3.6D,
            TextColor = ParsePdfColor(textBox.Color),
            FontSize = textBox.FontSize,
            Font = MapFont(textBox.FontName),
            Align = MapAlign(textBox.Paragraphs.FirstOrDefault()?.Alignment),
            VerticalAlign = MapTextVerticalAlign(textBox.TextVerticalAlignment)
        };
    }

    private static void RenderPicture(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointPicture picture, double x, double y, double width, double height, int slideNumber, PowerPointPdfSaveOptions options) {
        try {
            PdfCore.PdfImageStyle style = new();
            PptCore.PowerPointPictureCrop crop = picture.GetCrop();
            if (crop.HasCrop) {
                style.SourceCrop = new PdfCore.PdfImageSourceCrop(crop.Left, crop.Top, crop.Right, crop.Bottom);
            } else {
                style.Fit = options.PictureFit;
            }

            Uri? hyperlink = picture.ClickHyperlink;
            string? linkUri = null;
            if (hyperlink != null && hyperlink.IsAbsoluteUri) {
                linkUri = hyperlink.AbsoluteUri;
            } else if (hyperlink != null) {
                AddWarning(options, slideNumber, "relative-picture-hyperlink", "Skipped a relative PowerPoint picture hyperlink because PDF URI annotations require absolute targets.");
            }

            byte[] imageBytes = picture.GetImageBytes();
            AddPowerPointPictureAspectRatioDiagnostic(options, slideNumber, imageBytes, crop, style.Fit, x, y, width, height);

            canvas.Image(
                imageBytes,
                x,
                y,
                width,
                height,
                style: style,
                linkUri: linkUri,
                linkContents: linkUri != null ? picture.AltText : null,
                alternativeText: picture.AltText,
                rotationAngle: picture.Rotation ?? 0D,
                horizontalFlip: picture.HorizontalFlip == true,
                verticalFlip: picture.VerticalFlip == true);
        } catch (Exception ex) {
            AddLayoutWarning(
                options,
                slideNumber,
                "unsupported-picture",
                "Skipped a PowerPoint picture because it could not be embedded as a PDF image: " + ex.Message,
                PdfCore.PdfLayoutDiagnosticKind.SkippedContent,
                "PowerPointPicture",
                "The PowerPoint picture could not be embedded as a PDF image.",
                x,
                y,
                width,
                height);
        }
    }

    private static void RenderAutoShape(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointAutoShape autoShape, double x, double y, double width, double height, int slideNumber, PowerPointPdfSaveOptions options) {
        OfficeShape? shape = CreateOfficeShape(autoShape.ShapeType, autoShape, width, height);
        if (shape == null) {
            AddWarning(options, slideNumber, "unsupported-auto-shape", "Skipped unsupported PowerPoint auto-shape type '" + GetShapePresetName(autoShape.ShapeType) + "'.");
            return;
        }

        ApplyShapeStyle(autoShape, shape);
        canvas.Shape(shape, x, y, rotationAngle: autoShape.Rotation ?? 0D);
    }

    private static bool RenderTextBoxGeometry(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointTextBox textBox, double x, double y, double width, double height) {
        ShapeTypeValues? type = textBox.ShapeType;
        if (type == ShapeTypeValues.Rectangle) {
            return false;
        }

        OfficeShape? shape = CreateOfficeShape(type, textBox, width, height);
        if (shape == null) {
            return false;
        }

        ApplyShapeStyle(textBox, shape);
        canvas.Shape(shape, x, y, rotationAngle: textBox.Rotation ?? 0D);
        return true;
    }

    private static OfficeShape? CreateOfficeShape(ShapeTypeValues? type, PptCore.PowerPointShape source, double width, double height) {
        return type.HasValue &&
               OfficeShapePresets.TryCreate(GetShapePresetName(type), width, height, source.HorizontalFlip == true, source.VerticalFlip == true, out OfficeShape? shape)
            ? shape
            : null;
    }

    private static string? GetShapePresetName(ShapeTypeValues? type) {
        if (!type.HasValue) {
            return null;
        }

        ShapeTypeValues value = type.Value;
        if (value == ShapeTypeValues.Rectangle) return "rect";
        if (value == ShapeTypeValues.RoundRectangle) return "roundRect";
        if (value == ShapeTypeValues.Ellipse) return "ellipse";
        if (value == ShapeTypeValues.Line || value == ShapeTypeValues.StraightConnector1) return "line";
        if (value == ShapeTypeValues.Triangle) return "triangle";
        if (value == ShapeTypeValues.RightTriangle) return "rtTriangle";
        if (value == ShapeTypeValues.Diamond) return "diamond";
        if (value == ShapeTypeValues.Parallelogram) return "parallelogram";
        if (value == ShapeTypeValues.Trapezoid) return "trapezoid";
        if (value == ShapeTypeValues.Pentagon) return "pentagon";
        if (value == ShapeTypeValues.Hexagon) return "hexagon";
        if (value == ShapeTypeValues.Octagon) return "octagon";
        if (value == ShapeTypeValues.Plus) return "plus";
        if (value == ShapeTypeValues.Chevron) return "chevron";
        if (value == ShapeTypeValues.RightArrow) return "rightArrow";
        if (value == ShapeTypeValues.LeftArrow) return "leftArrow";
        if (value == ShapeTypeValues.UpArrow) return "upArrow";
        if (value == ShapeTypeValues.DownArrow) return "downArrow";
        if (value == ShapeTypeValues.Star5) return "star5";
        return value.ToString();
    }

    private static void ApplyShapeStyle(PptCore.PowerPointShape source, OfficeShape target) {
        target.FillColor = source.FillTransparency == 100 ? null : ParseOfficeColor(source.FillColor);
        if (source.FillTransparency.HasValue && source.FillTransparency.Value > 0 && source.FillTransparency.Value < 100) {
            target.FillOpacity = 1D - source.FillTransparency.Value / 100D;
        }

        target.StrokeColor = ParseOfficeColor(source.OutlineColor);
        target.StrokeWidth = source.OutlineWidthPoints ?? (target.StrokeColor.HasValue ? 1D : 0D);
        target.StrokeDashStyle = MapDash(source.OutlineDash);
    }

    private static bool TryGetShapeBox(PptCore.PowerPointShape shape, int slideNumber, double pageWidth, double pageHeight, PowerPointPdfSaveOptions options, bool warnInvalidBounds, out double x, out double y, out double width, out double height) {
        if (!shape.TryGetBoundsPoints(out x, out y, out width, out height)) {
            x = 0D;
            y = 0D;
            width = 0D;
            height = 0D;
        }

        if ((width <= 0D || height <= 0D) &&
            shape.OwnerSlide != null &&
            shape.ShapePlaceholderType.HasValue) {
            PptCore.PowerPointLayoutBox? layoutBounds = shape.OwnerSlide.GetLayoutPlaceholderBounds(shape.ShapePlaceholderType.Value, shape.ShapePlaceholderIndex);
            if (layoutBounds.HasValue) {
                x = layoutBounds.Value.LeftPoints;
                y = layoutBounds.Value.TopPoints;
                width = layoutBounds.Value.WidthPoints;
                height = layoutBounds.Value.HeightPoints;
            }
        }

        bool isLineShape = IsLineShape(shape);
        bool hasRenderableSize = isLineShape
            ? width >= 0D && height >= 0D && (width > 0D || height > 0D)
            : width > 0D && height > 0D;
        if (hasRenderableSize && IntersectsPage(x, y, width, height, pageWidth, pageHeight)) {
            return true;
        }

        if (warnInvalidBounds) {
            AddWarning(options, slideNumber, "invalid-shape-bounds", "Skipped a PowerPoint shape with non-positive or off-slide PDF bounds.");
        }

        return false;
    }

    private static bool IsLineShape(PptCore.PowerPointShape shape) {
        return shape is PptCore.PowerPointAutoShape autoShape &&
               (autoShape.ShapeType == ShapeTypeValues.Line || autoShape.ShapeType == ShapeTypeValues.StraightConnector1);
    }

    private static bool IntersectsPage(double x, double y, double width, double height, double pageWidth, double pageHeight) =>
        x + width > 0D && y + height > 0D && x < pageWidth && y < pageHeight;

    private static bool NeedsSlideClip(double x, double y, double width, double height, double pageWidth, double pageHeight) =>
        x < 0D || y < 0D || x + width > pageWidth || y + height > pageHeight;

    private static bool TryGetVisibleSlideBox(double x, double y, double width, double height, double pageWidth, double pageHeight, out double clipX, out double clipY, out double clipWidth, out double clipHeight) {
        clipX = Math.Max(0D, x);
        clipY = Math.Max(0D, y);
        double clipRight = Math.Min(pageWidth, x + width);
        double clipBottom = Math.Min(pageHeight, y + height);
        clipWidth = clipRight - clipX;
        clipHeight = clipBottom - clipY;
        return clipWidth > 0D && clipHeight > 0D;
    }

    private static PdfCore.PdfAlign MapAlign(TextAlignmentTypeValues? alignment) {
        if (alignment == TextAlignmentTypeValues.Center) {
            return PdfCore.PdfAlign.Center;
        }

        if (alignment == TextAlignmentTypeValues.Right) {
            return PdfCore.PdfAlign.Right;
        }

        return PdfCore.PdfAlign.Left;
    }

    private static PdfCore.PdfColumnAlign? MapColumnAlign(TextAlignmentTypeValues? alignment) {
        if (alignment == TextAlignmentTypeValues.Center) {
            return PdfCore.PdfColumnAlign.Center;
        }

        if (alignment == TextAlignmentTypeValues.Right) {
            return PdfCore.PdfColumnAlign.Right;
        }

        if (alignment == TextAlignmentTypeValues.Left) {
            return PdfCore.PdfColumnAlign.Left;
        }

        return null;
    }

    private static PdfCore.PdfCellVerticalAlign? MapVerticalAlign(TextAnchoringTypeValues? alignment) {
        if (alignment == TextAnchoringTypeValues.Center) {
            return PdfCore.PdfCellVerticalAlign.Middle;
        }

        if (alignment == TextAnchoringTypeValues.Bottom) {
            return PdfCore.PdfCellVerticalAlign.Bottom;
        }

        if (alignment == TextAnchoringTypeValues.Top) {
            return PdfCore.PdfCellVerticalAlign.Top;
        }

        return null;
    }

    private static PdfCore.PdfVerticalAlign MapTextVerticalAlign(TextAnchoringTypeValues? alignment) {
        if (alignment == TextAnchoringTypeValues.Center) {
            return PdfCore.PdfVerticalAlign.Middle;
        }

        if (alignment == TextAnchoringTypeValues.Bottom) {
            return PdfCore.PdfVerticalAlign.Bottom;
        }

        return PdfCore.PdfVerticalAlign.Top;
    }

    private static PdfCore.PdfStandardFont? MapFont(string? fontName) {
        if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(fontName, out PdfCore.PdfStandardFont font)) {
            return font;
        }

        return string.IsNullOrWhiteSpace(fontName)
            ? null
            : PdfCore.PdfStandardFont.Helvetica;
    }

    private static OfficeStrokeDashStyle MapDash(PresetLineDashValues? dash) {
        if (!dash.HasValue) {
            return OfficeStrokeDashStyle.Solid;
        }

        string value = dash.Value.ToString();
        if (value.IndexOf("Dot", StringComparison.OrdinalIgnoreCase) >= 0 &&
            value.IndexOf("Dash", StringComparison.OrdinalIgnoreCase) >= 0) {
            return OfficeStrokeDashStyle.DashDot;
        }

        if (value.IndexOf("Dot", StringComparison.OrdinalIgnoreCase) >= 0) {
            return OfficeStrokeDashStyle.Dot;
        }

        if (value.IndexOf("Dash", StringComparison.OrdinalIgnoreCase) >= 0) {
            return OfficeStrokeDashStyle.Dash;
        }

        return OfficeStrokeDashStyle.Solid;
    }

    private static OfficeLinearGradient? CreateLinearGradient(string startColor, string endColor, double angleDegrees) {
        OfficeColor? start = ParseOfficeColor(startColor);
        OfficeColor? end = ParseOfficeColor(endColor);
        if (!start.HasValue || !end.HasValue) {
            return null;
        }

        return OfficeLinearGradient.FromAngle(start.Value, end.Value, angleDegrees);
    }

    private static PdfCore.PdfColor? ParsePdfColor(string? value) {
        OfficeColor? color = ParseOfficeColor(value);
        return color.HasValue ? PdfCore.PdfColor.FromOfficeColorOrNull(color.Value) : null;
    }

    private static OfficeColor? ParseOfficeColor(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        string normalized = (value ?? string.Empty).Trim();
        if (normalized.StartsWith("#", StringComparison.Ordinal)) {
            normalized = normalized.Substring(1);
        }

        return OfficeColor.TryParseHex(normalized, out OfficeColor color) ? color : (OfficeColor?)null;
    }

    private static void AddWarning(PowerPointPdfSaveOptions options, int slideNumber, string code, string message) {
        var warning = new PowerPointPdfExportWarning(slideNumber, code, message);
        options.Warnings.Add(warning);
        options.Report.Add(warning.ToConversionWarning());
    }

    private static void AddWarning(PowerPointPdfSaveOptions options, int slideNumber, string code, string message, PdfCore.PdfLayoutDiagnostic layoutDiagnostic) {
        var warning = new PowerPointPdfExportWarning(slideNumber, code, message, layoutDiagnostic);
        options.Warnings.Add(warning);
        options.Report.Add(warning.ToConversionWarning());
    }
}

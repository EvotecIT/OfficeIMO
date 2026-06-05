using System;
using System.Collections.Generic;
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
        if (presentation == null) {
            throw new ArgumentNullException(nameof(presentation));
        }

        options ??= new PowerPointPdfSaveOptions();
        options.ResetExportState();
        PdfCore.PdfOptions pdfOptions = CreatePdfOptions(presentation, options);
        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create(pdfOptions);

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
    /// Converts a PowerPoint presentation to PDF bytes.
    /// </summary>
    public static byte[] SaveAsPdf(this PptCore.PowerPointPresentation presentation, PowerPointPdfSaveOptions? options = null) {
        return presentation.ToPdfDocument(options).ToBytes();
    }

    /// <summary>
    /// Saves a PowerPoint presentation as a PDF file.
    /// </summary>
    public static void SaveAsPdf(this PptCore.PowerPointPresentation presentation, string path, PowerPointPdfSaveOptions? options = null) {
        presentation.ToPdfDocument(options).Save(path);
    }

    /// <summary>
    /// Attempts to save a PowerPoint presentation as a PDF file and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this PptCore.PowerPointPresentation presentation, string path, PowerPointPdfSaveOptions? options = null) {
        try {
            return presentation.ToPdfDocument(options).TrySave(path);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(path, ex);
        }
    }

    /// <summary>
    /// Writes a PowerPoint presentation as PDF to a stream.
    /// </summary>
    public static void SaveAsPdf(this PptCore.PowerPointPresentation presentation, Stream stream, PowerPointPdfSaveOptions? options = null) {
        presentation.ToPdfDocument(options).Save(stream);
    }

    /// <summary>
    /// Attempts to write a PowerPoint presentation as PDF to a stream and returns output diagnostics instead of throwing.
    /// </summary>
    public static PdfCore.PdfSaveResult TrySaveAsPdf(this PptCore.PowerPointPresentation presentation, Stream stream, PowerPointPdfSaveOptions? options = null) {
        try {
            return presentation.ToPdfDocument(options).TrySave(stream);
        } catch (Exception ex) {
            return PdfCore.PdfSaveResult.FromFailure(outputPath: null, ex);
        }
    }

    private static PdfCore.PdfOptions CreatePdfOptions(PptCore.PowerPointPresentation presentation, PowerPointPdfSaveOptions options) {
        PdfCore.PdfOptions pdfOptions = options.PdfOptions?.Clone() ?? new PdfCore.PdfOptions();
        pdfOptions.PageWidth = presentation.SlideSize.WidthPoints;
        pdfOptions.PageHeight = presentation.SlideSize.HeightPoints;
        pdfOptions.Margins = PdfCore.PageMargins.Uniform(0);
        RegisterPresentationFonts(pdfOptions, presentation, options);
        return pdfOptions;
    }

    private static void RegisterPresentationFonts(PdfCore.PdfOptions pdfOptions, PptCore.PowerPointPresentation presentation, PowerPointPdfSaveOptions options) {
        var registeredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyList<PptCore.PowerPointSlide> slides = presentation.Slides;
        for (int slideIndex = 0; slideIndex < slides.Count; slideIndex++) {
            PptCore.PowerPointSlide slide = slides[slideIndex];
            if (!options.IncludeHiddenSlides && slide.Hidden) {
                continue;
            }

            RegisterPresentationShapesFonts(slide.GetInheritedShapesForExport(), pdfOptions, registeredFamilies);
            RegisterPresentationShapesFonts(slide.Shapes, pdfOptions, registeredFamilies);
        }
    }

    private static void RegisterPresentationShapesFonts(IReadOnlyList<PptCore.PowerPointShape> shapes, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies) {
        foreach (PptCore.PowerPointShape shape in shapes) {
            RegisterPresentationShapeFonts(shape, pdfOptions, registeredFamilies);
        }
    }

    private static void RegisterPresentationShapeFonts(PptCore.PowerPointShape shape, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies) {
        if (shape.Hidden) {
            return;
        }

        if (shape is PptCore.PowerPointTextBox textBox) {
            RegisterPresentationTextBoxFonts(textBox, pdfOptions, registeredFamilies);
            return;
        }

        if (shape is PptCore.PowerPointTable table) {
            RegisterPresentationTableFonts(table, pdfOptions, registeredFamilies);
            return;
        }

        if (shape is PptCore.PowerPointGroupShape groupShape && groupShape.OwnerSlide != null) {
            RegisterPresentationShapesFonts(groupShape.OwnerSlide.GetGroupChildren(groupShape), pdfOptions, registeredFamilies);
        }
    }

    private static void RegisterPresentationTextBoxFonts(PptCore.PowerPointTextBox textBox, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies) {
        RegisterPresentationFontCandidate(textBox.FontName, pdfOptions, registeredFamilies);
        foreach (PptCore.PowerPointParagraph paragraph in textBox.Paragraphs) {
            foreach (PptCore.PowerPointTextRun run in paragraph.Runs) {
                RegisterPresentationFontCandidate(run.FontName, pdfOptions, registeredFamilies);
            }
        }
    }

    private static void RegisterPresentationTableFonts(PptCore.PowerPointTable table, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies) {
        for (int row = 0; row < table.Rows; row++) {
            for (int column = 0; column < table.Columns; column++) {
                PptCore.PowerPointTableCell cell = table.GetCell(row, column);
                if (!cell.IsMergedCell) {
                    RegisterPresentationFontCandidate(cell.FontName, pdfOptions, registeredFamilies);
                }
            }
        }
    }

    private static void RegisterPresentationFontCandidate(string? familyName, PdfCore.PdfOptions pdfOptions, HashSet<string> registeredFamilies) {
        if (string.IsNullOrWhiteSpace(familyName)) {
            return;
        }

        string trimmedFamilyName = familyName!.Trim();
        if (!registeredFamilies.Add(trimmedFamilyName)) {
            return;
        }

        if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(trimmedFamilyName, out PdfCore.PdfStandardFont standardFont)) {
            pdfOptions.RegisterOfficeFontFamily(
                trimmedFamilyName,
                PdfCore.PdfStandardFontMapper.GetFontFamily(standardFont));
        }
    }

    private static void RenderSlide(PdfCore.PdfDocument pdf, PptCore.PowerPointSlide slide, int slideNumber, double pageWidth, double pageHeight, PowerPointPdfSaveOptions options) {
        pdf.Canvas(canvas => {
            RenderSlideBackground(canvas, slide, slideNumber, pageWidth, pageHeight, options);

            RenderShapes(canvas, slide.GetInheritedShapesForExport(), slideNumber, pageWidth, pageHeight, options, warnInvalidBounds: false);
            RenderShapes(canvas, slide.Shapes, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds: true);
        });
    }

    private static void RenderShapes(PdfCore.PdfPageCanvas canvas, IReadOnlyList<PptCore.PowerPointShape> shapes, int slideNumber, double pageWidth, double pageHeight, PowerPointPdfSaveOptions options, bool warnInvalidBounds) {
        foreach (PptCore.PowerPointShape shape in shapes) {
            if (shape.Hidden) {
                continue;
            }

            if (!TryGetShapeBox(shape, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, out double x, out double y, out double width, out double height)) {
                continue;
            }

            Action<PdfCore.PdfPageCanvas> render = target => RenderShapeContent(target, shape, x, y, width, height, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds);
            if (TryGetVisibleSlideBox(x, y, width, height, pageWidth, pageHeight, out double clipX, out double clipY, out double clipWidth, out double clipHeight) &&
                NeedsSlideClip(x, y, width, height, pageWidth, pageHeight)) {
                canvas.Clip(clipX, clipY, clipWidth, clipHeight, render);
            } else {
                render(canvas);
            }
        }
    }

    private static void RenderShapeContent(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointShape shape, double x, double y, double width, double height, int slideNumber, double pageWidth, double pageHeight, PowerPointPdfSaveOptions options, bool warnInvalidBounds) {
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
                RenderGroupShape(canvas, groupShape, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds);
            } else {
                AddWarning(options, slideNumber, "unsupported-shape", "Skipped a PowerPoint group shape because its owning slide context could not be resolved.");
            }
            return;
        }

        AddWarning(options, slideNumber, "unsupported-shape", "Skipped unsupported PowerPoint shape content type '" + shape.ShapeContentType + "'.");
    }

    private static void RenderGroupShape(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointGroupShape groupShape, int slideNumber, double pageWidth, double pageHeight, PowerPointPdfSaveOptions options, bool warnInvalidBounds) {
        IReadOnlyList<PptCore.PowerPointShape> children = groupShape.OwnerSlide!.GetGroupChildren(groupShape);
        foreach (PptCore.PowerPointShape child in children) {
            if (child.Hidden) {
                continue;
            }

            if (!TryGetShapeBox(child, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, out double x, out double y, out double width, out double height)) {
                continue;
            }

            MapGroupChildBox(groupShape, ref x, ref y, ref width, ref height);
            Action<PdfCore.PdfPageCanvas> render = target => RenderShapeContent(target, child, x, y, width, height, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds);
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
        canvas.TextBox(runs, x, y, width, height, style, textBox.Rotation ?? 0D);
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
        for (int index = 0; index < paragraphRuns.Count && cursorY < textY + textHeight; index++) {
            PptCore.PowerPointParagraph paragraph = textBox.Paragraphs.Count > index ? textBox.Paragraphs[index] : textBox.Paragraphs.Last();
            double paragraphHeight = Math.Min(paragraphHeights[index], textY + textHeight - cursorY);
            var paragraphStyle = CreateTransparentParagraphStyle(style, paragraph);
            canvas.TextBox(paragraphRuns[index], textX, cursorY, textWidth, Math.Max(1D, paragraphHeight), paragraphStyle);
            cursorY += paragraphHeight;
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
            }

            Uri? hyperlink = picture.ClickHyperlink;
            string? linkUri = null;
            if (hyperlink != null && hyperlink.IsAbsoluteUri) {
                linkUri = hyperlink.AbsoluteUri;
            } else if (hyperlink != null) {
                AddWarning(options, slideNumber, "relative-picture-hyperlink", "Skipped a relative PowerPoint picture hyperlink because PDF URI annotations require absolute targets.");
            }

            canvas.Image(
                picture.GetImageBytes(),
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
            AddWarning(options, slideNumber, "unsupported-picture", "Skipped a PowerPoint picture because it could not be embedded as a PDF image: " + ex.Message);
        }
    }

    private static void RenderTable(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointTable table, double x, double y, double width, double height, int slideNumber, PowerPointPdfSaveOptions options) {
        if (table.Rows == 0 || table.Columns == 0) {
            AddWarning(options, slideNumber, "empty-table", "Skipped an empty PowerPoint table.");
            return;
        }

        var rows = new List<PdfCore.PdfTableCell[]>();
        for (int rowIndex = 0; rowIndex < table.Rows; rowIndex++) {
            var pdfCells = new List<PdfCore.PdfTableCell>();
            for (int columnIndex = 0; columnIndex < table.Columns; columnIndex++) {
                PptCore.PowerPointTableCell cell = table.GetCell(rowIndex, columnIndex);
                if (cell.IsMergedCell) {
                    continue;
                }

                pdfCells.Add(CreatePdfTableCell(cell));
            }

            rows.Add(pdfCells.ToArray());
        }

        PdfCore.PdfTableStyle style = CreateTableStyle(table);
        try {
            canvas.Table(rows, x, y, width, height, style, table.Rotation ?? 0D);
        } catch (Exception ex) {
            AddWarning(options, slideNumber, "unsupported-table", "Skipped a PowerPoint table because it could not be rendered as a PDF table: " + ex.Message);
        }
    }

    private static PdfCore.PdfTableCell CreatePdfTableCell(PptCore.PowerPointTableCell cell) {
        (int rowSpan, int columnSpan) = cell.Merge;
        return PdfCore.PdfTableCell.Merge(CreatePdfTableCellRuns(cell), Math.Max(1, columnSpan), Math.Max(1, rowSpan));
    }

    private static IReadOnlyList<PdfCore.TextRun> CreatePdfTableCellRuns(PptCore.PowerPointTableCell cell) {
        var runs = new List<PdfCore.TextRun>();
        A.TextBody? textBody = cell.Cell.TextBody;
        if (textBody != null) {
            bool hasParagraph = false;
            foreach (A.Paragraph paragraph in textBody.Elements<A.Paragraph>()) {
                if (hasParagraph) {
                    runs.Add(PdfCore.TextRun.LineBreak());
                }

                AppendPdfTableCellParagraphRuns(runs, paragraph, cell);
                hasParagraph = true;
            }
        }

        if (runs.Count == 0) {
            runs.Add(CreatePdfTableCellTextRun(cell, cell.Text ?? string.Empty));
        }

        return runs;
    }

    private static void AppendPdfTableCellParagraphRuns(List<PdfCore.TextRun> runs, A.Paragraph paragraph, PptCore.PowerPointTableCell cell) {
        foreach (OpenXmlElement child in paragraph.ChildElements) {
            switch (child) {
                case A.Run run:
                    foreach (A.Text text in run.Elements<A.Text>()) {
                        runs.Add(CreatePdfTableCellTextRun(cell, text.Text ?? string.Empty));
                    }

                    break;
                case A.Break:
                    runs.Add(PdfCore.TextRun.LineBreak());
                    break;
                case A.Field field:
                    string fieldText = field.Text?.Text ?? field.InnerText ?? string.Empty;
                    if (!string.IsNullOrEmpty(fieldText)) {
                        runs.Add(CreatePdfTableCellTextRun(cell, fieldText));
                    }

                    break;
            }
        }
    }

    private static PdfCore.TextRun CreatePdfTableCellTextRun(PptCore.PowerPointTableCell cell, string text) {
        return new PdfCore.TextRun(
            text,
            bold: cell.Bold,
            italic: cell.Italic,
            color: ParsePdfColor(cell.Color),
            fontSize: cell.FontSize,
            font: MapFont(cell.FontName));
    }

    private static PdfCore.PdfTableStyle CreateTableStyle(PptCore.PowerPointTable table) {
        var style = PdfCore.TableStyles.Light().Clone();
        style.HeaderRowCount = table.FirstRow ? 1 : 0;
        style.FooterRowCount = table.LastRow ? 1 : 0;
        style.RowStripeFill = table.BandedRows ? style.RowStripeFill : null;
        style.ColumnWidthPoints = CreateColumnWidths(table, table.WidthPoints);
        style.RowMinHeights = CreateRowHeights(table, table.HeightPoints);
        style.CellFills = CreateTableCellFills(table);
        style.CellPaddings = CreateTableCellPaddings(table);
        style.CellAlignments = CreateTableCellAlignments(table);
        style.CellVerticalAlignments = CreateTableCellVerticalAlignments(table);
        style.CellBorders = CreateTableCellBorders(table);
        return style;
    }

    private static List<double?> CreateColumnWidths(PptCore.PowerPointTable table, double tableWidth) {
        var widths = new List<double?>(table.Columns);
        double fallbackWidth = table.Columns > 0 ? tableWidth / table.Columns : tableWidth;
        for (int column = 0; column < table.Columns; column++) {
            double width = table.GetColumnWidthPoints(column);
            widths.Add(width > 0D ? width : fallbackWidth);
        }

        return widths;
    }

    private static List<double?> CreateRowHeights(PptCore.PowerPointTable table, double tableHeight) {
        var heights = new List<double?>(table.Rows);
        double fallbackHeight = table.Rows > 0 ? tableHeight / table.Rows : tableHeight;
        for (int row = 0; row < table.Rows; row++) {
            double height = table.GetRowHeightPoints(row);
            heights.Add(height > 0D ? height : fallbackHeight);
        }

        return heights;
    }

    private static Dictionary<(int Row, int Column), PdfCore.PdfColor>? CreateTableCellFills(PptCore.PowerPointTable table) {
        var fills = new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
        ForEachTableAnchorCell(table, (row, column, cell) => {
            PdfCore.PdfColor? fill = ParsePdfColor(cell.FillColor);
            if (fill.HasValue) {
                fills[(row, column)] = fill.Value;
            }
        });

        return fills.Count == 0 ? null : fills;
    }

    private static Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>? CreateTableCellPaddings(PptCore.PowerPointTable table) {
        var paddings = new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>();
        ForEachTableAnchorCell(table, (row, column, cell) => {
            if (cell.PaddingLeftPoints.HasValue || cell.PaddingRightPoints.HasValue || cell.PaddingTopPoints.HasValue || cell.PaddingBottomPoints.HasValue) {
                paddings[(row, column)] = new PdfCore.PdfCellPadding {
                    Left = cell.PaddingLeftPoints,
                    Right = cell.PaddingRightPoints,
                    Top = cell.PaddingTopPoints,
                    Bottom = cell.PaddingBottomPoints
                };
            }
        });

        return paddings.Count == 0 ? null : paddings;
    }

    private static Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>? CreateTableCellAlignments(PptCore.PowerPointTable table) {
        var alignments = new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>();
        ForEachTableAnchorCell(table, (row, column, cell) => {
            PdfCore.PdfColumnAlign? align = MapColumnAlign(cell.HorizontalAlignment);
            if (align.HasValue) {
                alignments[(row, column)] = align.Value;
            }
        });

        return alignments.Count == 0 ? null : alignments;
    }

    private static Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>? CreateTableCellVerticalAlignments(PptCore.PowerPointTable table) {
        var alignments = new Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>();
        ForEachTableAnchorCell(table, (row, column, cell) => {
            PdfCore.PdfCellVerticalAlign? align = MapVerticalAlign(cell.VerticalAlignment);
            if (align.HasValue) {
                alignments[(row, column)] = align.Value;
            }
        });

        return alignments.Count == 0 ? null : alignments;
    }

    private static Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? CreateTableCellBorders(PptCore.PowerPointTable table) {
        var borders = new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>();
        ForEachTableAnchorCell(table, (row, column, cell) => {
            PdfCore.PdfColor? borderColor = ParsePdfColor(cell.BorderColor);
            if (borderColor.HasValue) {
                borders[(row, column)] = new PdfCore.PdfCellBorder {
                    Color = borderColor.Value,
                    Width = 0.75D
                };
            }
        });

        return borders.Count == 0 ? null : borders;
    }

    private static void ForEachTableAnchorCell(PptCore.PowerPointTable table, Action<int, int, PptCore.PowerPointTableCell> action) {
        for (int row = 0; row < table.Rows; row++) {
            for (int column = 0; column < table.Columns; column++) {
                PptCore.PowerPointTableCell cell = table.GetCell(row, column);
                if (!cell.IsMergedCell) {
                    action(row, column, cell);
                }
            }
        }
    }

    private static void RenderAutoShape(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointAutoShape autoShape, double x, double y, double width, double height, int slideNumber, PowerPointPdfSaveOptions options) {
        OfficeShape? shape = CreateOfficeShape(autoShape.ShapeType, autoShape, width, height);
        if (shape == null) {
            AddWarning(options, slideNumber, "unsupported-auto-shape", "Skipped unsupported PowerPoint auto-shape type '" + autoShape.ShapeType + "'.");
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
        if (type == ShapeTypeValues.Rectangle) {
            return OfficeShape.Rectangle(width, height);
        }

        if (type == ShapeTypeValues.RoundRectangle) {
            return OfficeShape.RoundedRectangle(width, height, Math.Min(width, height) * 0.18D);
        }

        if (type == ShapeTypeValues.Ellipse) {
            return OfficeShape.Ellipse(width, height);
        }

        if (type == ShapeTypeValues.Line) {
            double startX = source.HorizontalFlip == true ? width : 0D;
            double endX = source.HorizontalFlip == true ? 0D : width;
            double startY = source.VerticalFlip == true ? height : 0D;
            double endY = source.VerticalFlip == true ? 0D : height;
            return OfficeShape.Line(startX, startY, endX, endY);
        }

        return null;
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

        bool isLineShape = shape is PptCore.PowerPointAutoShape autoShape && autoShape.ShapeType == ShapeTypeValues.Line;
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
        return PdfCore.PdfStandardFontMapper.TryMapFontFamily(fontName, out PdfCore.PdfStandardFont font)
            ? font
            : null;
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

        double radians = angleDegrees * Math.PI / 180D;
        double dx = Math.Cos(radians);
        double dy = Math.Sin(radians);
        double startX = Clamp01(0.5D - dx / 2D);
        double startY = Clamp01(0.5D - dy / 2D);
        double endX = Clamp01(0.5D + dx / 2D);
        double endY = Clamp01(0.5D + dy / 2D);

        if (startX.Equals(endX) && startY.Equals(endY)) {
            return OfficeLinearGradient.Horizontal(start.Value, end.Value);
        }

        return new OfficeLinearGradient(
            startX,
            startY,
            endX,
            endY,
            new OfficeGradientStop(0D, start.Value),
            new OfficeGradientStop(1D, end.Value));
    }

    private static double Clamp01(double value) {
        if (value < 0D) {
            return 0D;
        }

        if (value > 1D) {
            return 1D;
        }

        return value;
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
        options.Warnings.Add(new PowerPointPdfExportWarning(slideNumber, code, message));
    }
}

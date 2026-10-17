using System;
using System.Globalization;
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

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderSlide(PdfCore.PdfDocument pdf, PptCore.PowerPointSlide slide, int slideNumber, double pageWidth, double pageHeight, PowerPointToPdfOptions options) {
        pdf.Canvas(canvas => {
            canvas.NamedDestination(PptCore.PowerPointHyperlinkResolver.GetSlideDestinationName(slideNumber), 0, 0);
            RenderSlideBackground(canvas, slide, slideNumber, pageWidth, pageHeight, options);

            RenderShapes(canvas, slide.GetInheritedShapesForExport(), slideNumber, pageWidth, pageHeight, options, warnInvalidBounds: false, groupDepth: 0);
            RenderShapes(canvas, slide.Shapes, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds: true, groupDepth: 0);
        });
    }

    private static void RenderNotesPages(PdfCore.PdfDocument pdf, PptCore.PowerPointPresentation presentation,
        double pageWidth, double pageHeight, PowerPointToPdfOptions options) {
        List<(PptCore.PowerPointSlide Slide, int Number)> slides = GetVisibleSlides(presentation, options);
        if (slides.Count == 0) {
            RenderEmptySlide(pdf, pageWidth, pageHeight);
            return;
        }
        for (int index = 0; index < slides.Count; index++) {
            options.CancellationToken.ThrowIfCancellationRequested();
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
                RenderSpeakerNotes(canvas, string.IsNullOrWhiteSpace(notes) ? "No speaker notes." : notes,
                    margin, margin + slideHeight + 48D, availableWidth,
                    Math.Max(36D, pageHeight - margin * 2D - slideHeight - 48D), 12D,
                    PdfCore.PdfColor.FromRgb(35, 42, 52), item.Number, options);
            });
        }
    }

    private static void RenderHandoutPages(PdfCore.PdfDocument pdf, PptCore.PowerPointPresentation presentation,
        double pageWidth, double pageHeight, PowerPointToPdfOptions options) {
        List<(PptCore.PowerPointSlide Slide, int Number)> slides = GetVisibleSlides(presentation, options);
        if (slides.Count == 0) {
            RenderEmptySlide(pdf, pageWidth, pageHeight);
            return;
        }
        int perPage = options.HandoutSlidesPerPage;
        for (int pageStart = 0, pageIndex = 0; pageStart < slides.Count; pageStart += perPage, pageIndex++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (pageIndex > 0) pdf.PageBreak();
            List<(PptCore.PowerPointSlide Slide, int Number)> pageSlides = slides.Skip(pageStart)
                .Take(perPage).ToList();
            pdf.Canvas(canvas => RenderHandoutPage(canvas, presentation, pageSlides, pageWidth, pageHeight, options));
        }
    }

    private static void RenderHandoutPage(PdfCore.PdfPageCanvas canvas,
        PptCore.PowerPointPresentation presentation,
        IReadOnlyList<(PptCore.PowerPointSlide Slide, int Number)> slides,
        double pageWidth, double pageHeight, PowerPointToPdfOptions options) {
        const double margin = 36D;
        const double gutter = 18D;
        if (options.HandoutSlidesPerPage == 3) {
            double rowHeight = (pageHeight - margin * 2D - gutter * 2D) / 3D;
            double thumbColumnWidth = (pageWidth - margin * 2D - gutter) * 0.46D;
            for (int index = 0; index < slides.Count; index++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                double top = margin + index * (rowHeight + gutter);
                double thumbHeight = Math.Min(rowHeight - 18D, thumbColumnWidth *
                    presentation.SlideSize.HeightPoints / presentation.SlideSize.WidthPoints);
                double thumbWidth = thumbHeight * presentation.SlideSize.WidthPoints /
                    presentation.SlideSize.HeightPoints;
                RenderSnapshotThumbnail(canvas, slides[index].Slide, slides[index].Number,
                    margin + (thumbColumnWidth - thumbWidth) / 2D, top, thumbWidth, thumbHeight, options);
                string notes = string.Empty;
                if (options.IncludeSpeakerNotes) slides[index].Slide.Notes.TryGetText(out notes);
                string noteText = string.IsNullOrWhiteSpace(notes)
                    ? "Slide " + slides[index].Number.ToString(CultureInfo.InvariantCulture) +
                      Environment.NewLine + "________________________________" + Environment.NewLine +
                      "________________________________" + Environment.NewLine + "________________________________"
                    : "Slide " + slides[index].Number.ToString(CultureInfo.InvariantCulture) +
                      Environment.NewLine + notes;
                RenderSpeakerNotes(canvas, noteText, margin + thumbColumnWidth + gutter, top,
                    pageWidth - margin * 2D - thumbColumnWidth - gutter, rowHeight,
                    9D, PdfCore.PdfColor.FromRgb(55, 62, 72), slides[index].Number, options);
            }
            return;
        }

        (int columns, int rows) = GetHandoutGrid(options.HandoutSlidesPerPage);
        double cellWidth = (pageWidth - margin * 2D - gutter * (columns - 1)) / columns;
        double cellHeight = (pageHeight - margin * 2D - gutter * (rows - 1)) / rows;
        for (int index = 0; index < slides.Count; index++) {
            options.CancellationToken.ThrowIfCancellationRequested();
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
                RenderSpeakerNotes(canvas, notes, left, top + thumbHeight + labelHeight + 2D, cellWidth, notesHeight,
                    7D, PdfCore.PdfColor.FromRgb(55, 62, 72), slides[index].Number, options);
            }
        }
    }

    private static void RenderSnapshotThumbnail(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointSlide slide,
        int slideNumber, double x, double y, double width, double height, PowerPointToPdfOptions options) {
        var size = slide.OwnerPresentation!.SlideSize;
        canvas.NamedDestination(PptCore.PowerPointHyperlinkResolver.GetSlideDestinationName(slideNumber), x, y);
        // Draw the canonical slide directly under a thumbnail transform. Reading
        // an image snapshot here would discard PDF links and text semantics.
        canvas.Clip(x, y, width, height, clipped =>
            clipped.Effect(new OfficeTransform(width / size.WidthPoints, 0D, 0D, height / size.HeightPoints, x, y), 1D, target => {
                RenderSlideBackground(target, slide, slideNumber, size.WidthPoints, size.HeightPoints, options);
                RenderShapes(target, slide.GetInheritedShapesForExport(), slideNumber, size.WidthPoints, size.HeightPoints,
                    options, warnInvalidBounds: false, groupDepth: 0);
                RenderShapes(target, slide.Shapes, slideNumber, size.WidthPoints, size.HeightPoints,
                    options, warnInvalidBounds: true, groupDepth: 0);
            }));
    }

    private static List<(PptCore.PowerPointSlide Slide, int Number)> GetVisibleSlides(
        PptCore.PowerPointPresentation presentation, PowerPointToPdfOptions options) {
        var slides = new List<(PptCore.PowerPointSlide, int)>();
        for (int index = 0; index < presentation.Slides.Count; index++) {
            PptCore.PowerPointSlide slide = presentation.Slides[index];
            if (!options.IncludesSlide(slide)) continue;
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

    private static void RenderEmptySlide(PdfCore.PdfDocument pdf, double pageWidth, double pageHeight) {
        pdf.Canvas(canvas => RenderFallbackBackground(canvas, pageWidth, pageHeight));
    }

    private static void RenderSlideBackground(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointSlide slide, int slideNumber, double pageWidth, double pageHeight, PowerPointToPdfOptions options) {
        if (!options.IncludeSlideBackgrounds) {
            RenderFallbackBackground(canvas, pageWidth, pageHeight);
            return;
        }

        PptCore.PowerPointSlideBackground background;
        try {
            background = slide.GetBackground(() => GetImageReadLimit(options));
            options.EmbeddedImageBytesRead += background.ImageByteLength;
        } catch (Exception exception) when (exception is InvalidOperationException ||
            exception is InvalidDataException invalid && OfficeIMO.Core.Internal.OfficeStreamReader.IsSizeLimitException(invalid)) {
            AddWarning(options, slideNumber, "embedded-image-limit", exception.Message, lossKind: OfficeConversionLossKind.Omission);
            RenderFallbackBackground(canvas, pageWidth, pageHeight);
            return;
        }
        if (background.Kind == PptCore.PowerPointSlideBackgroundKind.SolidColor && !string.IsNullOrWhiteSpace(background.Color)) {
            OfficeColor? color = ParseOfficeColor(background.Color);
            if (color.HasValue) {
                RenderBackgroundShape(canvas, pageWidth, pageHeight, color.Value, fillGradient: null);
                return;
            }

            AddWarning(options, slideNumber, "invalid-slide-background-color", "Used a white fallback because the PowerPoint slide background color could not be parsed.", lossKind: OfficeConversionLossKind.Omission);
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
                AddWarning(options, slideNumber, "unsupported-slide-background-image", "Used a white fallback because the PowerPoint slide background image could not be embedded as a PDF image: " + ex.Message, lossKind: OfficeConversionLossKind.Omission);
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

            AddWarning(options, slideNumber, "invalid-slide-background-gradient", "Used a white fallback because the PowerPoint slide background gradient colors could not be parsed.", lossKind: OfficeConversionLossKind.Omission);
        } else if (background.Kind == PptCore.PowerPointSlideBackgroundKind.Unsupported) {
            AddWarning(options, slideNumber, "unsupported-slide-background", "Used a white fallback because " + (background.UnsupportedReason ?? "the PowerPoint slide background is not supported."), lossKind: OfficeConversionLossKind.Omission);
        }

        RenderFallbackBackground(canvas, pageWidth, pageHeight);
    }

    private static void RenderFallbackBackground(PdfCore.PdfPageCanvas canvas, double pageWidth, double pageHeight) {
        RenderBackgroundShape(canvas, pageWidth, pageHeight, OfficeColor.Transparent, fillGradient: null);
    }

    private static void RenderBackgroundShape(PdfCore.PdfPageCanvas canvas, double pageWidth, double pageHeight, OfficeColor fillColor, OfficeLinearGradient? fillGradient) {
        OfficeShape background = OfficeShape.Rectangle(pageWidth, pageHeight);
        background.FillColor = fillColor;
        background.FillGradient = fillGradient?.Clone();
        background.StrokeWidth = 0D;
        canvas.Shape(background, 0, 0);
    }

    private static string ApplyPowerPointDisplayCase(
        string text,
        PptCore.PowerPointCapitalization? capitalization,
        string? language,
        PowerPointToPdfOptions options,
        int slideNumber) {
        if (capitalization is not (PptCore.PowerPointCapitalization.AllCaps or PptCore.PowerPointCapitalization.SmallCaps)) {
            return text;
        }

        if (capitalization == PptCore.PowerPointCapitalization.SmallCaps) {
            AddWarning(options, slideNumber, "small-caps-approximation", "Rendered PowerPoint small caps as uppercase because PDF text runs do not provide native small-cap glyph substitution.");
        }

        return OfficeTextCaseTransformer.Apply(text, OfficeTextCase.Uppercase, ResolvePowerPointRunCulture(language));
    }

    private static CultureInfo ResolvePowerPointRunCulture(string? language) {
        if (string.IsNullOrWhiteSpace(language)) return CultureInfo.InvariantCulture;
        try {
            return CultureInfo.GetCultureInfo(language);
        } catch (CultureNotFoundException) {
            return CultureInfo.InvariantCulture;
        }
    }

    private static OfficeTextDecorationStyle MapPowerPointUnderline(TextUnderlineValues? style) {
        if (!style.HasValue || style.Value == TextUnderlineValues.None) return OfficeTextDecorationStyle.None;
        if (style.Value == TextUnderlineValues.Double || style.Value == TextUnderlineValues.WavyDouble) return OfficeTextDecorationStyle.Double;
        if (style.Value == TextUnderlineValues.Dotted || style.Value == TextUnderlineValues.HeavyDotted) return OfficeTextDecorationStyle.Dotted;
        if (style.Value == TextUnderlineValues.Wavy || style.Value == TextUnderlineValues.WavyHeavy) return OfficeTextDecorationStyle.Wavy;
        if (style.Value == TextUnderlineValues.Dash || style.Value == TextUnderlineValues.DashHeavy ||
            style.Value == TextUnderlineValues.DashLong || style.Value == TextUnderlineValues.DashLongHeavy ||
            style.Value == TextUnderlineValues.DotDash || style.Value == TextUnderlineValues.DotDashHeavy ||
            style.Value == TextUnderlineValues.DotDotDash || style.Value == TextUnderlineValues.DotDotDashHeavy) {
            return OfficeTextDecorationStyle.Dashed;
        }

        return OfficeTextDecorationStyle.Single;
    }

    private static PdfCore.PdfTextBaseline MapPowerPointBaseline(
        double? baselinePercent,
        PowerPointToPdfOptions options,
        int slideNumber) {
        if (baselinePercent.HasValue && baselinePercent.Value != 0D &&
            Math.Abs(baselinePercent.Value - 30D) > 0.0001D &&
            Math.Abs(baselinePercent.Value + 25D) > 0.0001D &&
            !options.Warnings.Any(warning => warning.SlideNumber == slideNumber && warning.Code == "baseline-percent-approximation")) {
            AddWarning(options, slideNumber, "baseline-percent-approximation",
                "Rendered a native PowerPoint baseline percentage using the PDF superscript or subscript geometry; the exact authored percentage is not representable.");
        }

        return baselinePercent switch {
            > 0D => PdfCore.PdfTextBaseline.Superscript,
            < 0D => PdfCore.PdfTextBaseline.Subscript,
            _ => PdfCore.PdfTextBaseline.Normal
        };
    }

    private static void RenderPicture(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointPicture picture, double x, double y, double width, double height, int slideNumber, PowerPointToPdfOptions options) {
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
            }
            string? destination = ResolveSlideLinkDestination(hyperlink, slideNumber, options, "picture");

            byte[] imageBytes = picture.GetImageBytes(GetImageReadLimit(options));
            options.EmbeddedImageBytesRead += imageBytes.Length;
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
                rotationAngle: -(picture.Rotation ?? 0D),
                horizontalFlip: picture.HorizontalFlip == true,
                verticalFlip: picture.VerticalFlip == true);
            if (destination != null) RenderSlideLinkFrame(canvas, picture, destination, x, y, width, height);
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

    private static void RenderAutoShape(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointAutoShape autoShape, double x, double y, double width, double height, int slideNumber, PowerPointToPdfOptions options) {
        if (!autoShape.ShapeType.HasValue) {
            if (PptCore.PowerPointSlideImageRenderer
                    .TryCreateCustomGeometryDrawing(autoShape, width, height,
                        out OfficeDrawing? drawing, out string? reason)
                && drawing != null) {
                canvas.Drawing(drawing, x, y, width, height,
                    style: new PdfCore.PdfDrawingStyle {
                        AlternativeText = string.IsNullOrWhiteSpace(autoShape.Name)
                            ? "PowerPoint custom geometry"
                            : autoShape.Name
                    });
            } else {
                AddLayoutWarning(options, slideNumber,
                    "unsupported-custom-geometry",
                    "Skipped a PowerPoint custom-geometry shape: " + reason,
                    PdfCore.PdfLayoutDiagnosticKind.SkippedContent,
                    "PowerPointCustomGeometry",
                    "The custom geometry could not be projected through the shared Drawing renderer.",
                    x, y, width, height);
            }
            return;
        }
        OfficeShape? shape = CreateOfficeShape(autoShape.ShapeType.ToOpenXml(), autoShape, width, height);
        if (shape == null) {
            AddWarning(options, slideNumber, "unsupported-auto-shape", "Skipped unsupported PowerPoint auto-shape type '" + GetShapePresetName(autoShape.ShapeType.ToOpenXml()) + "'.", lossKind: OfficeConversionLossKind.Omission);
            return;
        }

        ApplyShapeAppearance(autoShape, shape, slideNumber, options);
        canvas.Shape(shape, x, y, rotationAngle: autoShape.Rotation ?? 0D);
    }

    private static bool RenderTextBoxGeometry(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointTextBox textBox, double x, double y, double width, double height, int slideNumber, PowerPointToPdfOptions options) {
        ShapeTypeValues? type = textBox.ShapeType.ToOpenXml();
        if (type == ShapeTypeValues.Rectangle) {
            return false;
        }

        OfficeShape? shape = CreateOfficeShape(type, textBox, width, height);
        if (shape == null) {
            return false;
        }

        ApplyShapeAppearance(textBox, shape, slideNumber, options);
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
        if (!type.HasValue) return null;
        return type.Value == ShapeTypeValues.Line || type.Value == ShapeTypeValues.StraightConnector1
            ? "line" : ((DocumentFormat.OpenXml.IEnumValue)type.Value).Value;
    }

    private static void ApplyShapeAppearance(PptCore.PowerPointShape source, OfficeShape target,
        int slideNumber, PowerPointToPdfOptions options) {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        PptCore.PowerPointSlideImageRenderer.ApplyShapeAppearanceForExport(target, source, diagnostics);
        foreach (OfficeImageExportDiagnostic diagnostic in diagnostics) {
            AddWarning(options, slideNumber, diagnostic.Code, diagnostic.Message, lossKind: diagnostic.LossKind);
        }
    }

    private static bool TryGetShapeBox(PptCore.PowerPointShape shape, int slideNumber, double pageWidth, double pageHeight, PowerPointToPdfOptions options, bool warnInvalidBounds, out double x, out double y, out double width, out double height, bool applyPageCulling = true) {
        shape.TryGetExportBoundsPoints(out x, out y, out width, out height);

        bool isLineShape = IsLineShape(shape);
        bool hasRenderableSize = isLineShape
            ? width >= 0D && height >= 0D && (width > 0D || height > 0D)
            : width > 0D && height > 0D;
        // A group's final visible bounds depend on all descendant transforms, not its raw box.
        if (hasRenderableSize && (!applyPageCulling || shape is PptCore.PowerPointGroupShape || IntersectsPage(x, y, width, height, pageWidth, pageHeight))) {
            return true;
        }

        if (warnInvalidBounds) {
            AddWarning(options, slideNumber, "invalid-shape-bounds", "Skipped a PowerPoint shape with non-positive or off-slide PDF bounds.", lossKind: OfficeConversionLossKind.Omission);
        }

        return false;
    }

    private static bool IsLineShape(PptCore.PowerPointShape shape) {
        OfficePresetShapeType? type = shape switch {
            PptCore.PowerPointAutoShape autoShape => autoShape.ShapeType,
            PptCore.PowerPointConnectionShape connection => connection.ShapeType,
            _ => null
        };
        return type == OfficePresetShapeType.Line || type == OfficePresetShapeType.StraightConnector1;
    }

    private static bool IntersectsPage(double x, double y, double width, double height, double pageWidth, double pageHeight) =>
        x + width > 0D && y + height > 0D && x < pageWidth && y < pageHeight;



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

}

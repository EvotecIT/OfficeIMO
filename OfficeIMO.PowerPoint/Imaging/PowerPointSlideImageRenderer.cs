using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        internal static OfficeImageExportResult Render(PowerPointSlide slide, OfficeImageExportFormat format, PowerPointImageExportOptions options) {
            if (slide == null) {
                throw new ArgumentNullException(nameof(slide));
            }

            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            PowerPointSlideVisualSnapshot snapshot = CreateSnapshot(slide, options);
            OfficeDrawing drawing = snapshot.Drawing;

            if (format == OfficeImageExportFormat.Svg) {
                List<OfficeImageExportDiagnostic> diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
                var fallbackCodec = new OfficeRasterImageFallbackCodec(options.ImageCodec, diagnostics, "PowerPoint slide");
                byte[] svg = OfficeDrawingSvgExporter.ToSvgBytes(drawing, options.Scale, OfficeSvgSizeUnit.Pixel, fallbackCodec);
                return new OfficeImageExportResult(format, ScaledWidth(drawing, options), ScaledHeight(drawing, options), svg, "Slide", "PowerPoint slide", diagnostics);
            }

            if (format.IsRaster()) {
                List<OfficeImageExportDiagnostic> diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
                const string source = "PowerPoint slide";
                OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(
                    drawing.Width,
                    drawing.Height,
                    format,
                    options,
                    source);
                if (plan.Diagnostic != null) diagnostics.Add(plan.Diagnostic);
                var fallbackCodec = new OfficeRasterImageFallbackCodec(options.ImageCodec, diagnostics, source);
                OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
                    Scale = plan.Limit.Scale,
                    Background = options.BackgroundColor,
                    ImageCodec = fallbackCodec
                });
                byte[] bytes = OfficeRasterImageEncoder.Encode(image, format, options.RasterEncoding);
                return new OfficeImageExportResult(format, image.Width, image.Height, bytes, "Slide", source, diagnostics);
            }

            throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported image export format.");
        }

        internal static PowerPointSlideVisualSnapshot CreateSnapshot(PowerPointSlide slide, PowerPointImageExportOptions options) {
            List<OfficeImageExportDiagnostic> diagnostics = new List<OfficeImageExportDiagnostic>();
            OfficeDrawing drawing = CreateDrawing(slide, options, diagnostics);
            return new PowerPointSlideVisualSnapshot(drawing, diagnostics.AsReadOnly());
        }

        private static OfficeDrawing CreateDrawing(PowerPointSlide slide, PowerPointImageExportOptions options, List<OfficeImageExportDiagnostic> diagnostics) {
            (double width, double height) = GetSlideSizePoints(slide);
            OfficeDrawing drawing = new OfficeDrawing(width, height);
            AddBackgroundRectangle(drawing, options.BackgroundColor, fillGradient: null);

            if (options.IncludeSlideBackground) {
                AddResolvedBackground(slide, drawing, diagnostics);
            }

            if (options.IncludeSlideContent) {
                A.ColorScheme? colorScheme = GetSlideColorScheme(slide);
                AddSlideContent(slide.GetInheritedShapesForExport(), drawing, diagnostics,
                    PowerPointShapeBoundsMapping.Identity, colorScheme, options);
                AddSlideContent(slide.Shapes, drawing, diagnostics,
                    PowerPointShapeBoundsMapping.Identity, colorScheme, options);
            }

            return drawing;
        }

        private static void AddSlideContent(IEnumerable<PowerPointShape> shapes, OfficeDrawing drawing,
            List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping,
            A.ColorScheme? colorScheme, PowerPointImageExportOptions options) {
            foreach (PowerPointShape shape in shapes) {
                if (shape.Hidden && !options.IncludeHiddenShapes) {
                    continue;
                }
                if (!ShouldIncludeShape(shape, options)) {
                    continue;
                }

                if (shape is PowerPointGroupShape groupShape) {
                    AddGroupShape(drawing, groupShape, diagnostics, mapping, colorScheme, options);
                } else if (shape is PowerPointPicture picture) {
                    AddPicture(drawing, picture, diagnostics, mapping);
                } else if (shape is PowerPointTable table) {
                    AddTable(drawing, table, diagnostics, mapping, colorScheme);
                } else if (shape is PowerPointChart chart) {
                    AddChart(drawing, chart, diagnostics, mapping, colorScheme);
                } else if (shape is PowerPointTextBox textBox) {
                    AddTextBox(drawing, textBox, diagnostics, mapping, colorScheme);
                } else if (shape is PowerPointAutoShape autoShape) {
                    AddAutoShape(drawing, autoShape, diagnostics, mapping, colorScheme);
                } else if (shape is PowerPointConnectionShape connectionShape) {
                    AddAutoShape(drawing, connectionShape, diagnostics, mapping, colorScheme);
                } else if (HasUnsupportedTransform(shape)) {
                    AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint shape because rotated or flipped slide content is not yet projected through OfficeIMO.Drawing.");
                    continue;
                } else {
                    AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint shape type that is not yet projected through OfficeIMO.Drawing.");
                }
            }
        }

        private static A.ColorScheme? GetSlideColorScheme(PowerPointSlide slide) =>
            slide.SlidePart.ThemeOverridePart?.ThemeOverride?.ColorScheme
            ?? slide.SlidePart.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?
                .ColorScheme
            ?? slide.SlidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?
                .ThemeElements?.ColorScheme;

        private static A.FormatScheme? GetSlideFormatScheme(PowerPointSlide slide) =>
            slide.SlidePart.ThemeOverridePart?.ThemeOverride?.FormatScheme
            ?? slide.SlidePart.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?
                .FormatScheme
            ?? slide.SlidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?
                .ThemeElements?.FormatScheme;

        private static void AddGroupShape(OfficeDrawing drawing, PowerPointGroupShape groupShape,
            List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping,
            A.ColorScheme? colorScheme, PowerPointImageExportOptions options) {
            if (groupShape.OwnerSlide == null) {
                AddUnsupportedShapeDiagnostic(diagnostics, groupShape, "Skipped a PowerPoint group shape because its owning slide context could not be resolved.");
                return;
            }

            OfficeImageFrameTransform frameTransform = CreateGroupFrameTransform(groupShape, 0D, 0D, 0D, 0D);
            if (frameTransform.HasTransform) {
                if (!TryGetBounds(groupShape, drawing, diagnostics, mapping, out double left, out double top, out double width, out double height)) {
                    return;
                }

                var groupDrawing = new OfficeDrawing(Math.Max(1D, drawing.Width - left), Math.Max(1D, drawing.Height - top));
                PowerPointShapeBoundsMapping localChildMapping = CreateGroupLocalChildMapping(groupShape, mapping);
                AddSlideContent(groupShape.OwnerSlide.GetGroupChildren(groupShape), groupDrawing, diagnostics,
                    localChildMapping, colorScheme, options);

                try {
                    OfficeImageFrameTransform groupFrameTransform = CreateGroupFrameTransform(groupShape, left, top, width, height);
                    if (RequiresGroupClip(groupDrawing, width, height)) {
                        drawing.AddClippedDrawing(groupDrawing, left, top, OfficeClipPath.Rectangle(width, height), groupFrameTransform);
                    } else {
                        drawing.AddDrawing(groupDrawing, left, top, groupFrameTransform);
                    }
                } catch (ArgumentOutOfRangeException) {
                    AddUnsupportedShapeDiagnostic(diagnostics, groupShape, "Skipped a PowerPoint group shape because its transformed Drawing content would exceed the slide canvas.");
                }

                return;
            }

            PowerPointShapeBoundsMapping childMapping = CreateGroupChildMapping(groupShape, mapping);
            if (TryGetBounds(groupShape, drawing, diagnostics, mapping, out double clipLeft, out double clipTop, out double clipWidth, out double clipHeight)) {
                var groupDrawing = new OfficeDrawing(Math.Max(1D, drawing.Width - clipLeft), Math.Max(1D, drawing.Height - clipTop));
                PowerPointShapeBoundsMapping localChildMapping = CreateGroupLocalChildMapping(groupShape, mapping);
                AddSlideContent(groupShape.OwnerSlide.GetGroupChildren(groupShape), groupDrawing, diagnostics,
                    localChildMapping, colorScheme, options);
                if (RequiresGroupClip(groupDrawing, clipWidth, clipHeight)) {
                    drawing.AddClippedDrawing(groupDrawing, clipLeft, clipTop, OfficeClipPath.Rectangle(clipWidth, clipHeight));
                    return;
                }
            }

            AddSlideContent(groupShape.OwnerSlide.GetGroupChildren(groupShape), drawing, diagnostics,
                childMapping, colorScheme, options);
        }

        private static bool RequiresGroupClip(OfficeDrawing groupDrawing, double width, double height) {
            for (int i = 0; i < groupDrawing.Elements.Count; i++) {
                OfficeDrawingElement element = groupDrawing.Elements[i];
                if (element is OfficeDrawingShape shape) {
                    if (shape.X < 0D || shape.Y < 0D || shape.X + shape.Shape.Width > width || shape.Y + shape.Shape.Height > height) {
                        return true;
                    }
                } else if (element is OfficeDrawingText text) {
                    if (text.X < 0D || text.Y < 0D || text.X + text.Width > width || text.Y + text.Height > height) {
                        return true;
                    }
                } else if (element is OfficeDrawingRichText richText) {
                    if (richText.X < 0D || richText.Y < 0D || richText.X + richText.Width > width || richText.Y + richText.Height > height) {
                        return true;
                    }
                } else if (element is OfficeDrawingImage image) {
                    (double left, double top, double right, double bottom) = image.Projection.GetDestinationBounds();
                    if (left < 0D || top < 0D || right > width || bottom > height) {
                        return true;
                    }
                } else if (element is OfficeDrawingGroup group) {
                    if (group.X < 0D || group.Y < 0D || group.X + group.ClipPath.Width > width || group.Y + group.ClipPath.Height > height) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static void AddPicture(OfficeDrawing drawing, PowerPointPicture picture, List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping) {
            if (!TryGetBounds(picture, drawing, diagnostics, mapping, out double left, out double top, out double width, out double height)) {
                return;
            }

            byte[] bytes;
            try {
                bytes = picture.GetImageBytes();
            } catch (InvalidOperationException) {
                AddUnsupportedShapeDiagnostic(diagnostics, picture, "Skipped a PowerPoint picture because its embedded image bytes could not be read.");
                return;
            }

            if (bytes.Length == 0) {
                AddUnsupportedShapeDiagnostic(diagnostics, picture, "Skipped a PowerPoint picture because its embedded image bytes are empty.");
                return;
            }

            PowerPointPictureCrop crop = picture.GetCrop();
            OfficeImageProjection projection = CreateImageProjection(
                left,
                top,
                width,
                height,
                crop.Left,
                crop.Top,
                crop.Right,
                crop.Bottom,
                picture.Rotation ?? 0D,
                picture.HorizontalFlip == true,
                picture.VerticalFlip == true);
            TryAddImage(drawing, bytes, picture.ContentType, projection, DescribeShape(picture), diagnostics, picture);
        }

        private static void AddAutoShape(OfficeDrawing drawing, PowerPointShape shape, List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping, A.ColorScheme? colorScheme) {
            string? presetName = GetAutoShapePresetName(shape);
            if (IsZeroThicknessLinePreset(presetName) && TryAddZeroThicknessLine(drawing, shape, diagnostics, mapping, presetName, colorScheme)) {
                return;
            }

            if (IsBentConnectorPreset(presetName) && TryAddBentConnector(drawing, shape, diagnostics, mapping, colorScheme)) {
                return;
            }

            if (IsCurvedConnectorPreset(presetName) && TryAddCurvedConnector(drawing, shape, diagnostics, mapping, colorScheme)) {
                return;
            }

            if (!TryGetBounds(shape, drawing, diagnostics, mapping, out double left, out double top, out double width, out double height)) {
                return;
            }

            if (TryAddCustomGeometryShape(drawing, shape, left, top, width, height, diagnostics, mapping, colorScheme)) {
                return;
            }

            if (!OfficeShapePresets.TryCreate(presetName, width, height, out OfficeShape? drawingShape) || drawingShape == null) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint auto shape geometry that is not yet projected through OfficeIMO.Drawing.");
                return;
            }

            ApplyShapeStyle(drawingShape, shape, colorScheme, mapping, diagnostics);
            ApplyShapeTransform(drawingShape, shape, width, height);
            drawing.AddShape(drawingShape, left, top);
        }

        private static string? GetAutoShapePresetName(PowerPointShape shape) {
            return GetOpenXmlShapeProperties(shape)?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText;
        }

        private static void AddTextBox(OfficeDrawing drawing, PowerPointTextBox textBox, List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping, A.ColorScheme? colorScheme) {
            string text = textBox.Text;
            bool hasVisibleFrame = HasVisibleFrame(textBox, colorScheme);
            if (string.IsNullOrEmpty(text) && !hasVisibleFrame) {
                return;
            }

            if (!TryGetBounds(textBox, drawing, diagnostics, mapping, out double left, out double top, out double width, out double height)) {
                return;
            }

            OfficeShape frame = CreateTextBoxFrame(textBox, width, height, diagnostics);
            ApplyShapeStyle(frame, textBox, colorScheme, mapping, diagnostics);
            ApplyShapeTransform(frame, textBox, width, height);
            if (hasVisibleFrame) {
                drawing.AddShape(frame, left, top);
            }

            if (string.IsNullOrEmpty(text)) {
                return;
            }

            double marginLeft = mapping.MapHorizontalLength(textBox.TextMarginLeftPoints ?? 0D);
            double marginTop = mapping.MapVerticalLength(textBox.TextMarginTopPoints ?? 0D);
            double marginRight = mapping.MapHorizontalLength(textBox.TextMarginRightPoints ?? 0D);
            double marginBottom = mapping.MapVerticalLength(textBox.TextMarginBottomPoints ?? 0D);
            double textWidth = width - marginLeft - marginRight;
            double textHeight = height - marginTop - marginBottom;
            if (textWidth <= 0D || textHeight <= 0D) {
                AddUnsupportedShapeDiagnostic(diagnostics, textBox, "Skipped PowerPoint text because the text margins leave no renderable drawing area.");
                return;
            }

            var padding = new OfficeTextPadding(marginLeft, marginTop, marginRight, marginBottom);
            PowerPointParagraph? firstParagraph = textBox.Paragraphs.Count > 0 ? textBox.Paragraphs[0] : null;
            OfficeTextAlignment alignment = MapTextAlignment(firstParagraph?.Alignment);
            OfficeTextParagraphIndent paragraphIndent = CreateParagraphIndent(firstParagraph, mapping);
            double rotation = textBox.Rotation ?? 0D;
            double rotationCenterX = left + (width / 2D);
            double rotationCenterY = top + (height / 2D);
            bool flipHorizontal = textBox.HorizontalFlip == true;
            bool flipVertical = textBox.VerticalFlip == true;
            if (TryAddTextBoxParagraphFlow(
                drawing,
                textBox,
                left,
                top,
                width,
                height,
                textWidth,
                textHeight,
                marginLeft,
                marginTop,
                rotation,
                rotationCenterX,
                rotationCenterY,
                flipHorizontal,
                flipVertical,
                mapping,
                colorScheme,
                diagnostics)) {
                return;
            }

            List<OfficeRichTextRun> richRuns = CreateRichTextRuns(textBox, colorScheme, mapping);
            if (ShouldRenderRichText(richRuns)) {
                drawing.AddRichText(
                    richRuns,
                    left,
                    top,
                    width,
                    height,
                    alignment,
                    verticalAlignment: MapTextVerticalAlignment(textBox.TextVerticalAlignment),
                    rotationDegrees: rotation,
                    rotationCenterX: rotationCenterX,
                    rotationCenterY: rotationCenterY,
                    wrapText: true,
                    flipHorizontal: flipHorizontal,
                    flipVertical: flipVertical,
                    padding: padding,
                    paragraphIndent: paragraphIndent);
                return;
            }

            drawing.AddText(
                text,
                left,
                top,
                width,
                height,
                CreateFont(textBox, mapping),
                ResolveTextBoxColor(textBox, colorScheme),
                alignment,
                rotationDegrees: rotation,
                rotationCenterX: rotationCenterX,
                rotationCenterY: rotationCenterY,
                verticalAlignment: MapTextVerticalAlignment(textBox.TextVerticalAlignment),
                wrapText: true,
                flipHorizontal: flipHorizontal,
                flipVertical: flipVertical,
                padding: padding,
                paragraphIndent: paragraphIndent);
        }

        private static OfficeShape CreateTextBoxFrame(
            PowerPointTextBox textBox,
            double width,
            double height,
            List<OfficeImageExportDiagnostic> diagnostics) {
            string? presetName = GetAutoShapePresetName(textBox);
            if (string.IsNullOrEmpty(presetName)
                || string.Equals(presetName, "rect", StringComparison.OrdinalIgnoreCase)
                || string.Equals(presetName, "rectangle", StringComparison.OrdinalIgnoreCase)) {
                return OfficeShape.Rectangle(width, height);
            }
            if (OfficeShapePresets.TryCreate(presetName, width, height,
                    out OfficeShape? preset) && preset != null) {
                return preset;
            }

            AddUnsupportedShapeDiagnostic(diagnostics, textBox,
                "Rendered the PowerPoint text content inside a rectangular frame because its preset frame geometry is not yet projected through OfficeIMO.Drawing.");
            return OfficeShape.Rectangle(width, height);
        }

        private static bool HasVisibleFrame(PowerPointShape source, A.ColorScheme? colorScheme) {
            return (source.FillTransparency != 100
                    && (HasShapeFillGradient(source)
                        || TryResolveShapeFillColor(source, colorScheme, out _))) ||
                TryResolveShapeOutlineColor(source, colorScheme, out _);
        }

        private static void ApplyShapeStyle(OfficeShape target, PowerPointShape source,
            A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping,
            List<OfficeImageExportDiagnostic> diagnostics) {
            if (source.FillTransparency != 100) {
                ShapeFillGradientProjection gradientProjection = ApplyShapeFillGradient(
                    target, source, colorScheme, mapping.HasTransformedAncestor);
                if (gradientProjection == ShapeFillGradientProjection.Unsupported) {
                    AddUnsupportedShapeDiagnostic(diagnostics, source,
                        "Skipped a PowerPoint shape gradient that cannot be represented faithfully by the shared Drawing renderer.");
                } else if (gradientProjection == ShapeFillGradientProjection.None
                    && TryResolveShapeFillColor(source, colorScheme, out OfficeColor fill)) {
                    target.FillColor = fill;
                }
            }

            if (TryResolveShapeOutlineColor(source, colorScheme, out OfficeColor stroke)) {
                target.StrokeColor = stroke;
                target.StrokeWidth = mapping.MapStrokeWidth(source.OutlineWidthPoints ?? 1D);
                target.StrokeDashStyle = MapDash(source.OutlineDash);
                ApplyStrokeLineCapAndJoin(target, source);
                ApplyLineMarkers(target, source);
            } else {
                target.StrokeWidth = 0D;
            }

            ApplyShapeEffects(target, source, colorScheme, mapping);
        }

        private static bool TryResolveShapeFillColor(PowerPointShape source, A.ColorScheme? colorScheme, out OfficeColor color) {
            color = default;
            DocumentFormat.OpenXml.Presentation.ShapeProperties? properties = GetOpenXmlShapeProperties(source);
            if (properties != null) {
                OfficeColor? resolvedColor = OfficeOpenXmlThemeColorResolver.ResolveColor(properties.GetFirstChild<A.SolidFill>(), colorScheme);
                if (resolvedColor.HasValue) {
                    color = resolvedColor.Value;
                    return true;
                }
            }

            return TryParseOfficeColor(source.FillColor, out color);
        }

        private static bool TryResolveShapeOutlineColor(PowerPointShape source, A.ColorScheme? colorScheme, out OfficeColor color) {
            color = default;
            DocumentFormat.OpenXml.Presentation.ShapeProperties? properties = GetOpenXmlShapeProperties(source);
            if (properties != null) {
                OfficeColor? resolvedColor = OfficeOpenXmlThemeColorResolver.ResolveColor(properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>(), colorScheme);
                if (resolvedColor.HasValue) {
                    color = resolvedColor.Value;
                    return true;
                }
            }

            return TryParseOfficeColor(source.OutlineColor, out color);
        }

        private static void ApplyStrokeLineCapAndJoin(OfficeShape target, PowerPointShape source) {
            A.Outline? outline = GetOpenXmlShapeProperties(source)?.GetFirstChild<A.Outline>();
            if (outline == null) {
                return;
            }

            target.StrokeLineCap = MapLineCap(outline.CapType?.Value);
            target.StrokeLineJoin = MapLineJoin(outline);
        }

        private static OfficeStrokeLineCap? MapLineCap(A.LineCapValues? cap) {
            if (!cap.HasValue) {
                return null;
            }

            if (cap.Value == A.LineCapValues.Round) {
                return OfficeStrokeLineCap.Round;
            }

            if (cap.Value == A.LineCapValues.Square) {
                return OfficeStrokeLineCap.Square;
            }

            if (cap.Value == A.LineCapValues.Flat) {
                return OfficeStrokeLineCap.Butt;
            }

            return null;
        }

        private static OfficeStrokeLineJoin? MapLineJoin(A.Outline outline) {
            if (outline.GetFirstChild<A.Round>() != null) {
                return OfficeStrokeLineJoin.Round;
            }

            if (outline.GetFirstChild<A.Bevel>() != null) {
                return OfficeStrokeLineJoin.Bevel;
            }

            if (outline.GetFirstChild<A.Miter>() != null) {
                return OfficeStrokeLineJoin.Miter;
            }

            return null;
        }

        private static void ApplyLineMarkers(OfficeShape target, PowerPointShape source) {
            if (target.Kind != OfficeShapeKind.Line && target.Kind != OfficeShapeKind.Path) {
                return;
            }

            A.Outline? outline = GetOpenXmlShapeProperties(source)?.GetFirstChild<A.Outline>();
            if (outline == null) {
                return;
            }

            A.HeadEnd? headEnd = outline.GetFirstChild<A.HeadEnd>();
            A.TailEnd? tailEnd = outline.GetFirstChild<A.TailEnd>();
            target.StrokeStartMarker = MapLineMarker(headEnd?.Type?.Value, headEnd?.Width?.Value, headEnd?.Length?.Value, target.StrokeWidth);
            target.StrokeEndMarker = MapLineMarker(tailEnd?.Type?.Value, tailEnd?.Width?.Value, tailEnd?.Length?.Value, target.StrokeWidth);
        }

        private static OfficeLineMarker? MapLineMarker(A.LineEndValues? type, A.LineEndWidthValues? width, A.LineEndLengthValues? length, double strokeWidth) {
            OfficeLineMarkerKind kind = MapLineMarkerKind(type);
            if (kind == OfficeLineMarkerKind.None) {
                return null;
            }

            double markerWidth = Math.Max(1D, strokeWidth * MapLineMarkerWidthFactor(width));
            double markerLength = Math.Max(1D, strokeWidth * MapLineMarkerLengthFactor(length));
            return new OfficeLineMarker(kind, markerWidth, markerLength);
        }

        private static DocumentFormat.OpenXml.Presentation.ShapeProperties? GetOpenXmlShapeProperties(PowerPointShape shape) {
            return shape.Element switch {
                DocumentFormat.OpenXml.Presentation.Shape openXmlShape => openXmlShape.ShapeProperties,
                DocumentFormat.OpenXml.Presentation.ConnectionShape connectionShape => connectionShape.ShapeProperties,
                _ => null
            };
        }

        private static OfficeLineMarkerKind MapLineMarkerKind(A.LineEndValues? type) {
            if (type == A.LineEndValues.Triangle) {
                return OfficeLineMarkerKind.Triangle;
            }

            if (type == A.LineEndValues.Stealth) {
                return OfficeLineMarkerKind.Stealth;
            }

            if (type == A.LineEndValues.Diamond) {
                return OfficeLineMarkerKind.Diamond;
            }

            if (type == A.LineEndValues.Oval) {
                return OfficeLineMarkerKind.Oval;
            }

            if (type == A.LineEndValues.Arrow) {
                return OfficeLineMarkerKind.Arrow;
            }

            return OfficeLineMarkerKind.None;
        }

        private static double MapLineMarkerWidthFactor(A.LineEndWidthValues? width) {
            if (width == A.LineEndWidthValues.Small) {
                return 3D;
            }

            if (width == A.LineEndWidthValues.Large) {
                return 6D;
            }

            return 4.5D;
        }

        private static double MapLineMarkerLengthFactor(A.LineEndLengthValues? length) {
            if (length == A.LineEndLengthValues.Small) {
                return 4D;
            }

            if (length == A.LineEndLengthValues.Large) {
                return 8D;
            }

            return 6D;
        }

        private static void ApplyShapeTransform(OfficeShape target, PowerPointShape source, double width, double height) {
            double rotation = source.Rotation ?? 0D;
            bool flipHorizontal = source.HorizontalFlip == true;
            bool flipVertical = source.VerticalFlip == true;
            var transform = new OfficeImageFrameTransform(
                rotation,
                width / 2D,
                height / 2D,
                flipHorizontal,
                flipVertical);
            if (transform.HasTransform) {
                target.Transform = transform.CreateDestinationTransform();
            }
        }

        private static OfficeFontInfo CreateFont(PowerPointTextBox textBox, PowerPointShapeBoundsMapping mapping) {
            OfficeFontStyle style = OfficeFontStyle.Regular;
            if (textBox.Bold) {
                style |= OfficeFontStyle.Bold;
            }

            if (textBox.Italic) {
                style |= OfficeFontStyle.Italic;
            }

            return new OfficeFontInfo(textBox.FontName ?? "Calibri", mapping.MapFontSize(textBox.FontSize ?? 18), style);
        }

        private static bool ShouldRenderRichText(IReadOnlyList<OfficeRichTextRun> richRuns) =>
            richRuns.Count > 1 ||
            (richRuns.Count == 1 && (richRuns[0].Underline || richRuns[0].Strikethrough || richRuns[0].BackgroundColor.HasValue));

        private static OfficeTextParagraphIndent CreateParagraphIndent(PowerPointParagraph? paragraph, PowerPointShapeBoundsMapping mapping) {
            if (paragraph == null) {
                return OfficeTextParagraphIndent.Empty;
            }

            double leftMargin = Math.Max(0D, mapping.MapHorizontalLength(paragraph.LeftMarginPoints ?? 0D));
            double firstLine = Math.Max(0D, leftMargin + mapping.MapHorizontalLength(paragraph.IndentPoints ?? 0D));
            return firstLine > 0D || leftMargin > 0D
                ? new OfficeTextParagraphIndent(firstLine, leftMargin)
                : OfficeTextParagraphIndent.Empty;
        }

        private static List<OfficeRichTextRun> CreateRichTextRuns(PowerPointTextBox textBox, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping) {
            IReadOnlyList<PowerPointParagraph> paragraphs = textBox.Paragraphs;
            var richRuns = new List<OfficeRichTextRun>();
            var numberingState = new Dictionary<int, int>();
            for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++) {
                PowerPointParagraph paragraph = paragraphs[paragraphIndex];
                IReadOnlyList<PowerPointTextRun> paragraphRuns = paragraph.Runs;
                PowerPointTextRun? firstRun = paragraphRuns.Count > 0 ? paragraphRuns[0] : null;
                string? marker = CreateParagraphMarker(paragraph, numberingState);
                bool paragraphHasVisibleText = !string.IsNullOrEmpty(marker) || paragraphRuns.Any(run => !string.IsNullOrEmpty(run.Text));
                if (!paragraphHasVisibleText) {
                    continue;
                }

                if (richRuns.Count > 0) {
                    richRuns.Add(CreateRichTextRun(Environment.NewLine, firstRun, textBox, paragraph, colorScheme, mapping));
                }

                if (!string.IsNullOrEmpty(marker)) {
                    richRuns.Add(CreateRichTextRun(marker!, firstRun, textBox, paragraph, colorScheme, mapping, markerRun: true));
                }

                for (int runIndex = 0; runIndex < paragraphRuns.Count; runIndex++) {
                    PowerPointTextRun run = paragraphRuns[runIndex];
                    string runText = run.Text;
                    if (string.IsNullOrEmpty(runText)) {
                        continue;
                    }

                    richRuns.Add(CreateRichTextRun(runText, run, textBox, paragraph, colorScheme, mapping));
                }
            }

            return richRuns;
        }

        private static string? CreateParagraphMarker(PowerPointParagraph paragraph, Dictionary<int, int> numberingState) {
            int level = paragraph.Level ?? 0;
            string? bullet = paragraph.BulletCharacter;
            if (!string.IsNullOrEmpty(bullet)) {
                return bullet + " ";
            }

            A.AutoNumberedBullet? numbered = paragraph.Paragraph.ParagraphProperties?.GetFirstChild<A.AutoNumberedBullet>();
            if (numbered == null) {
                return null;
            }

            int current = numbered.StartAt?.Value ?? (numberingState.TryGetValue(level, out int previous) ? previous + 1 : 1);
            numberingState[level] = current;
            return PowerPointNumberingFormatter.FormatMarker(current,
                numbered.Type?.Value) + " ";
        }

        private static OfficeRichTextRun CreateRichTextRun(string text, PowerPointTextRun? run, PowerPointTextBox textBox, PowerPointParagraph? paragraph, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping, bool markerRun = false) {
            OfficeColor color = ResolveTextRunColor(run, textBox, colorScheme);
            OfficeColor? backgroundColor = ResolveTextRunBackgroundColor(run, colorScheme);
            return new OfficeRichTextRun(
                text,
                mapping.MapFontSize(markerRun ? paragraph?.BulletSizePoints ?? run?.FontSize ?? textBox.FontSize ?? 18 : run?.FontSize ?? textBox.FontSize ?? 18),
                color,
                run?.Bold == true,
                run?.Italic == true,
                run?.Underline == true,
                markerRun ? paragraph?.BulletFontName ?? run?.FontName ?? textBox.FontName ?? "Calibri" : run?.FontName ?? textBox.FontName ?? "Calibri",
                run?.Strikethrough == true,
                backgroundColor);
        }

        private static OfficeColor ResolveTextRunColor(PowerPointTextRun? run, PowerPointTextBox textBox, A.ColorScheme? colorScheme) {
            OfficeColor? runColor = OfficeOpenXmlThemeColorResolver.ResolveColor(run?.Run.RunProperties?.GetFirstChild<A.SolidFill>(), colorScheme);
            if (runColor.HasValue) {
                return runColor.Value;
            }

            return ResolveTextBoxColor(textBox, colorScheme);
        }

        private static OfficeColor ResolveTextBoxColor(PowerPointTextBox textBox, A.ColorScheme? colorScheme) {
            A.Run? run = textBox.Paragraphs
                .SelectMany(paragraph => paragraph.Paragraph.Elements<A.Run>())
                .FirstOrDefault();
            OfficeColor? textBoxColor = OfficeOpenXmlThemeColorResolver.ResolveColor(run?.RunProperties?.GetFirstChild<A.SolidFill>(), colorScheme);
            return textBoxColor.HasValue
                ? textBoxColor.Value
                : OfficeColor.Black;
        }

        private static OfficeColor? ResolveTextRunBackgroundColor(PowerPointTextRun? run, A.ColorScheme? colorScheme) {
            return OfficeOpenXmlThemeColorResolver.ResolveColor(run?.Run.RunProperties?.GetFirstChild<A.Highlight>(), colorScheme);
        }

        private static OfficeTextAlignment MapTextAlignment(A.TextAlignmentTypeValues? alignment) {
            if (alignment == A.TextAlignmentTypeValues.Center) {
                return OfficeTextAlignment.Center;
            }

            if (alignment == A.TextAlignmentTypeValues.Right) {
                return OfficeTextAlignment.Right;
            }

            if (alignment == A.TextAlignmentTypeValues.Justified ||
                alignment == A.TextAlignmentTypeValues.Distributed ||
                alignment == A.TextAlignmentTypeValues.ThaiDistributed ||
                alignment == A.TextAlignmentTypeValues.JustifiedLow) {
                return OfficeTextAlignment.Justify;
            }

            return OfficeTextAlignment.Left;
        }

        private static OfficeTextVerticalAlignment MapTextVerticalAlignment(A.TextAnchoringTypeValues? alignment) {
            if (alignment == A.TextAnchoringTypeValues.Center) {
                return OfficeTextVerticalAlignment.Center;
            }

            if (alignment == A.TextAnchoringTypeValues.Bottom) {
                return OfficeTextVerticalAlignment.Bottom;
            }

            return OfficeTextVerticalAlignment.Top;
        }

        private static OfficeStrokeDashStyle MapDash(A.PresetLineDashValues? dash) {
            if (dash == A.PresetLineDashValues.LargeDashDotDot ||
                dash == A.PresetLineDashValues.SystemDashDotDot) {
                return OfficeStrokeDashStyle.DashDotDot;
            }

            if (dash == A.PresetLineDashValues.DashDot ||
                dash == A.PresetLineDashValues.LargeDashDot ||
                dash == A.PresetLineDashValues.SystemDashDot) {
                return OfficeStrokeDashStyle.DashDot;
            }

            if (dash == A.PresetLineDashValues.Dot ||
                dash == A.PresetLineDashValues.SystemDot) {
                return OfficeStrokeDashStyle.Dot;
            }

            if (dash == A.PresetLineDashValues.Dash ||
                dash == A.PresetLineDashValues.LargeDash ||
                dash == A.PresetLineDashValues.SystemDash) {
                return OfficeStrokeDashStyle.Dash;
            }

            return OfficeStrokeDashStyle.Solid;
        }

        private static bool TryGetBounds(PowerPointShape shape, OfficeDrawing drawing, List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping, out double left, out double top, out double width, out double height) {
            if (!shape.TryGetBoundsPoints(out left, out top, out width, out height)) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint shape because its bounds are outside the slide drawing canvas.");
                return false;
            }

            left = mapping.MapX(left);
            top = mapping.MapY(top);
            width = mapping.MapWidth(width);
            height = mapping.MapHeight(height);

            if (width <= 0D ||
                height <= 0D ||
                left < 0D ||
                top < 0D ||
                left + width > drawing.Width ||
                top + height > drawing.Height) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint shape because its bounds are outside the slide drawing canvas.");
                return false;
            }

            return true;
        }

        private static PowerPointShapeBoundsMapping CreateGroupChildMapping(PowerPointGroupShape groupShape, PowerPointShapeBoundsMapping parentMapping) {
            A.TransformGroup? transform = groupShape.GroupShape.GroupShapeProperties?.TransformGroup;
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
                return parentMapping;
            }

            double groupX = PowerPointUnits.ToPoints(groupXEmu.Value);
            double groupY = PowerPointUnits.ToPoints(groupYEmu.Value);
            double childX = PowerPointUnits.ToPoints(childXEmu.Value);
            double childY = PowerPointUnits.ToPoints(childYEmu.Value);
            double scaleX = groupWidthEmu.Value / (double)childWidthEmu.Value;
            double scaleY = groupHeightEmu.Value / (double)childHeightEmu.Value;
            PowerPointShapeBoundsMapping groupMapping = new PowerPointShapeBoundsMapping(
                groupX - (childX * scaleX),
                groupY - (childY * scaleY),
                scaleX,
                scaleY);
            return parentMapping.Compose(groupMapping);
        }

        private static PowerPointShapeBoundsMapping CreateGroupLocalChildMapping(PowerPointGroupShape groupShape, PowerPointShapeBoundsMapping parentMapping) {
            A.TransformGroup? transform = groupShape.GroupShape.GroupShapeProperties?.TransformGroup;
            bool hasTransformedAncestor = parentMapping.HasTransformedAncestor
                || CreateGroupFrameTransform(groupShape, 0D, 0D, 0D, 0D).HasTransform;
            long? groupWidthEmu = transform?.Extents?.Cx?.Value;
            long? groupHeightEmu = transform?.Extents?.Cy?.Value;
            long? childXEmu = transform?.ChildOffset?.X?.Value;
            long? childYEmu = transform?.ChildOffset?.Y?.Value;
            long? childWidthEmu = transform?.ChildExtents?.Cx?.Value;
            long? childHeightEmu = transform?.ChildExtents?.Cy?.Value;
            if (!groupWidthEmu.HasValue || !groupHeightEmu.HasValue ||
                !childXEmu.HasValue || !childYEmu.HasValue || !childWidthEmu.HasValue || !childHeightEmu.HasValue ||
                childWidthEmu.Value == 0L || childHeightEmu.Value == 0L) {
                return parentMapping.WithTransformedAncestor(hasTransformedAncestor);
            }

            double childX = PowerPointUnits.ToPoints(childXEmu.Value);
            double childY = PowerPointUnits.ToPoints(childYEmu.Value);
            double scaleX = groupWidthEmu.Value / (double)childWidthEmu.Value;
            double scaleY = groupHeightEmu.Value / (double)childHeightEmu.Value;
            return new PowerPointShapeBoundsMapping(
                -parentMapping.MapWidth(childX * scaleX),
                -parentMapping.MapHeight(childY * scaleY),
                parentMapping.MapWidth(scaleX),
                parentMapping.MapHeight(scaleY),
                hasTransformedAncestor);
        }

        private static OfficeImageFrameTransform CreateGroupFrameTransform(PowerPointGroupShape groupShape, double left, double top, double width, double height) {
            A.TransformGroup? transform = groupShape.GroupShape.GroupShapeProperties?.TransformGroup;
            double rotation = transform?.Rotation?.Value / 60000D ?? 0D;
            return new OfficeImageFrameTransform(
                rotation,
                left + (width / 2D),
                top + (height / 2D),
                transform?.HorizontalFlip?.Value == true,
                transform?.VerticalFlip?.Value == true);
        }

        private static bool HasUnsupportedTransform(PowerPointShape shape) {
            double rotation = shape.Rotation ?? 0D;
            return Math.Abs(rotation) > 0.000001D || shape.HorizontalFlip == true || shape.VerticalFlip == true;
        }

        private static OfficeImageProjection CreateImageProjection(
            double left,
            double top,
            double width,
            double height,
            double cropLeft,
            double cropTop,
            double cropRight,
            double cropBottom,
            double rotationDegrees,
            bool flipHorizontal,
            bool flipVertical) {
            return new OfficeImageProjection(
                new OfficeImagePlacement(left, top, width, height),
                OfficeImageSourceCrop.FromClampedFractions(cropLeft, cropTop, cropRight, cropBottom),
                rotationDegrees,
                left + (width / 2D),
                top + (height / 2D),
                flipHorizontal,
                flipVertical);
        }

        private static void TryAddImage(
            OfficeDrawing drawing,
            byte[] bytes,
            string? contentType,
            OfficeImageProjection projection,
            string source,
            List<OfficeImageExportDiagnostic> diagnostics,
            PowerPointShape? shape) {
            try {
                drawing.AddImage(bytes, contentType, projection, source);
            } catch (ArgumentOutOfRangeException) {
                if (shape == null) {
                    AddDiagnostic(diagnostics, "unsupported-powerpoint-image-bounds", "Skipped a PowerPoint image because its projected bounds are outside the slide drawing canvas.");
                } else {
                    AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint image because its projected bounds are outside the slide drawing canvas.");
                }
            }
        }

        private static void AddResolvedBackground(PowerPointSlide slide, OfficeDrawing drawing, List<OfficeImageExportDiagnostic> diagnostics) {
            PowerPointSlideBackground background = slide.GetBackground();
            switch (background.Kind) {
                case PowerPointSlideBackgroundKind.None:
                    return;
                case PowerPointSlideBackgroundKind.SolidColor:
                    if (TryParseOfficeColor(background.Color, out OfficeColor color)) {
                        AddBackgroundRectangle(drawing, color, fillGradient: null);
                    } else {
                        AddDiagnostic(diagnostics, "invalid-slide-background-color", "Used the configured fallback background because the PowerPoint slide background color could not be parsed.");
                    }
                    return;
                case PowerPointSlideBackgroundKind.LinearGradient:
                    if (TryCreateGradient(background, out OfficeLinearGradient? gradient)) {
                        AddBackgroundRectangle(drawing, gradient!.Stops[0].Color, gradient);
                    } else {
                        AddDiagnostic(diagnostics, "invalid-slide-background-gradient", "Used the configured fallback background because the PowerPoint slide background gradient colors could not be parsed.");
                    }
                    return;
                case PowerPointSlideBackgroundKind.Image:
                    AddBackgroundImage(drawing, background, diagnostics);
                    return;
                case PowerPointSlideBackgroundKind.Unsupported:
                    AddDiagnostic(diagnostics, "unsupported-slide-background", "Used the configured fallback background because " + (background.UnsupportedReason ?? "the PowerPoint slide background is not supported."));
                    return;
            }
        }

        private static void AddBackgroundImage(OfficeDrawing drawing, PowerPointSlideBackground background, List<OfficeImageExportDiagnostic> diagnostics) {
            byte[]? bytes = background.ImageBytes;
            if (bytes == null || bytes.Length == 0) {
                AddDiagnostic(diagnostics, "invalid-slide-background-image", "Used the configured fallback background because the PowerPoint slide background image bytes could not be read.");
                return;
            }

            OfficeImageProjection projection = CreateImageProjection(
                0D,
                0D,
                drawing.Width,
                drawing.Height,
                background.ImageCropLeft,
                background.ImageCropTop,
                background.ImageCropRight,
                background.ImageCropBottom,
                rotationDegrees: 0D,
                flipHorizontal: false,
                flipVertical: false);
            TryAddImage(drawing, bytes, background.ImageContentType, projection, "PowerPoint slide background image", diagnostics, null);
        }

        private static void AddBackgroundRectangle(OfficeDrawing drawing, OfficeColor fillColor, OfficeLinearGradient? fillGradient) {
            OfficeShape shape = OfficeShape.Rectangle(drawing.Width, drawing.Height);
            shape.FillColor = fillColor;
            shape.FillGradient = fillGradient?.Clone();
            shape.StrokeWidth = 0D;
            drawing.AddShape(shape, 0D, 0D);
        }

        private static bool TryCreateGradient(PowerPointSlideBackground background, out OfficeLinearGradient? gradient) {
            gradient = null;
            if (!TryParseOfficeColor(background.GradientStartColor, out OfficeColor start) ||
                !TryParseOfficeColor(background.GradientEndColor, out OfficeColor end)) {
                return false;
            }

            gradient = OfficeLinearGradient.FromAngle(start, end, background.GradientAngleDegrees ?? 0D);
            return true;
        }

        private static bool TryParseOfficeColor(string? value, out OfficeColor color) {
            color = default;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            string normalized = value!.Trim();
            if (normalized.StartsWith("#", StringComparison.Ordinal)) {
                normalized = normalized.Substring(1);
            }

            return OfficeColor.TryParseHex(normalized, out color);
        }

        private static (double Width, double Height) GetSlideSizePoints(PowerPointSlide slide) {
            long widthEmus = 12192000L;
            long heightEmus = 6858000L;
            if (slide.SlidePart.OpenXmlPackage is DocumentFormat.OpenXml.Packaging.PresentationDocument presentationDocument) {
                DocumentFormat.OpenXml.Presentation.SlideSize? size = presentationDocument.PresentationPart?.Presentation?.SlideSize;
                widthEmus = size?.Cx?.Value ?? widthEmus;
                heightEmus = size?.Cy?.Value ?? heightEmus;
            }

            return (PowerPointUnits.ToPoints(widthEmus), PowerPointUnits.ToPoints(heightEmus));
        }

        private static int ScaledWidth(OfficeDrawing drawing, PowerPointImageExportOptions options) =>
            Math.Max(1, (int)Math.Ceiling(drawing.Width * options.Scale));

        private static int ScaledHeight(OfficeDrawing drawing, PowerPointImageExportOptions options) =>
            Math.Max(1, (int)Math.Ceiling(drawing.Height * options.Scale));

        private static int UnscaledWidth(OfficeDrawing drawing) =>
            Math.Max(1, (int)Math.Ceiling(drawing.Width));

        private static int UnscaledHeight(OfficeDrawing drawing) =>
            Math.Max(1, (int)Math.Ceiling(drawing.Height));

        private static void AddDiagnostic(List<OfficeImageExportDiagnostic> diagnostics, string code, string message) {
            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                code,
                message,
                "PowerPoint slide"));
        }

        private static string DescribeShape(PowerPointShape shape) =>
            string.IsNullOrWhiteSpace(shape.Name) ? "PowerPoint shape" : "PowerPoint shape '" + shape.Name + "'";

        private static void AddUnsupportedShapeDiagnostic(List<OfficeImageExportDiagnostic> diagnostics, PowerPointShape shape, string message) {
            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                "unsupported-powerpoint-shape",
                message,
                DescribeShape(shape)));
        }

        private readonly struct PowerPointShapeBoundsMapping {
            internal static PowerPointShapeBoundsMapping Identity { get; } = new PowerPointShapeBoundsMapping(0D, 0D, 1D, 1D);

            private readonly double _offsetX;
            private readonly double _offsetY;
            private readonly double _scaleX;
            private readonly double _scaleY;
            private readonly bool _hasTransformedAncestor;

            internal PowerPointShapeBoundsMapping(double offsetX, double offsetY,
                double scaleX, double scaleY,
                bool hasTransformedAncestor = false) {
                _offsetX = offsetX;
                _offsetY = offsetY;
                _scaleX = scaleX;
                _scaleY = scaleY;
                _hasTransformedAncestor = hasTransformedAncestor;
            }

            internal bool HasTransformedAncestor => _hasTransformedAncestor;

            internal double MapX(double value) => _offsetX + (value * _scaleX);

            internal double MapY(double value) => _offsetY + (value * _scaleY);

            internal double MapWidth(double value) => value * _scaleX;

            internal double MapHeight(double value) => value * _scaleY;

            internal double MapHorizontalLength(double value) => value * Math.Abs(_scaleX);

            internal double MapVerticalLength(double value) => value * Math.Abs(_scaleY);

            internal double MapFontSize(double value) => value * Math.Abs(_scaleY);

            internal double MapStrokeWidth(double value) => value * Math.Sqrt(Math.Abs(_scaleX * _scaleY));

            internal PowerPointShapeBoundsMapping WithTransformedAncestor(bool value) =>
                new PowerPointShapeBoundsMapping(_offsetX, _offsetY, _scaleX,
                    _scaleY, _hasTransformedAncestor || value);

            internal PowerPointShapeBoundsMapping Compose(PowerPointShapeBoundsMapping child) =>
                new PowerPointShapeBoundsMapping(
                    _offsetX + (child._offsetX * _scaleX),
                    _offsetY + (child._offsetY * _scaleY),
                    _scaleX * child._scaleX,
                    _scaleY * child._scaleY,
                    _hasTransformedAncestor || child._hasTransformedAncestor);
        }
    }
}

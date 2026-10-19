using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static void AddDrawing(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderDrawing visual,
        RegisteredWebFonts webFonts, PdfImageResourceCache imageResources,
        PdfCore.PdfConversionReport conversionReport,
        CancellationToken cancellationToken,
        bool suppressLink) {
        OfficeDrawing source = visual.Drawing;
        double scaleX = visual.Width / source.Width;
        double scaleY = visual.Height / source.Height;
        if (TryAddRasterizedDrawingEffect(
                canvas,
                visual,
                source,
                Math.Max(scaleX, scaleY),
                imageResources,
                conversionReport,
                cancellationToken,
                suppressLink)) return;
        double originX = visual.X * PointsPerCssPixel;
        double originY = visual.Y * PointsPerCssPixel;
        bool fragmentLink = !suppressLink && IsFragmentLink(visual.LinkUri);
        OfficeTransform drawingToPage = OfficeTransform.Scale(scaleX * PointsPerCssPixel, scaleY * PointsPerCssPixel)
            .Then(OfficeTransform.Translate(originX, originY));
        OfficeTransform pageToDrawing = drawingToPage.Invert();

        void AddElements(PdfCore.PdfPageCanvas target, IReadOnlyList<OfficeDrawingElement> elements) {
            var shapeBatch = new OfficeDrawing(source.Width, source.Height);
            void FlushShapes() {
                if (shapeBatch.Elements.Count == 0) return;
                cancellationToken.ThrowIfCancellationRequested();
                target.DrawingForClippedRendering(
                    shapeBatch,
                    originX,
                    originY,
                    visual.Width * PointsPerCssPixel,
                    visual.Height * PointsPerCssPixel);
                shapeBatch = new OfficeDrawing(source.Width, source.Height);
            }

            foreach (OfficeDrawingElement element in elements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (element is OfficeDrawingShape shape) {
                    // Nested SVG clips, filters, and foreign-object viewports may retain
                    // deliberately out-of-bounds paint that is clipped by an ancestor group.
                    shapeBatch.AddShapeForClippedRendering(shape.Shape, shape.X, shape.Y);
                    continue;
                }
                if (element is OfficeDrawingGroup drawingGroup) {
                    FlushShapes();
                    OfficeTransform groupTransform = OfficeTransform.Translate(drawingGroup.X, drawingGroup.Y);
                    if (drawingGroup.FrameTransform.HasValue && drawingGroup.FrameTransform.Value.HasTransform) {
                        groupTransform = groupTransform.Then(drawingGroup.FrameTransform.Value.CreateDestinationTransform());
                    }
                    OfficeTransform pageGroupTransform = pageToDrawing
                        .Then(groupTransform)
                        .Then(drawingToPage);
                    OfficeDrawing nestedDrawing = drawingGroup.Drawing;
                    Action<PdfCore.PdfPageCanvas> addGroupPaint = groupTarget => groupTarget.Effect(pageGroupTransform, 1D, grouped => {
                        grouped.SourceClip(
                            originX,
                            originY,
                            drawingGroup.ClipPath.Scale(scaleX * PointsPerCssPixel, scaleY * PointsPerCssPixel),
                            clipped => {
                                OfficeTransform contentTransform = OfficeTransform.Translate(
                                    drawingGroup.ContentOffsetX,
                                    drawingGroup.ContentOffsetY);
                                if (contentTransform == OfficeTransform.Identity) {
                                    AddElements(clipped, nestedDrawing.Elements);
                                } else {
                                    OfficeTransform pageContentTransform = pageToDrawing
                                        .Then(contentTransform)
                                        .Then(drawingToPage);
                                    clipped.Effect(pageContentTransform, 1D,
                                        nested => AddElements(nested, nestedDrawing.Elements));
                                }
                            });
                    });
                    if (drawingGroup.ActualText == null) {
                        addGroupPaint(target);
                    } else {
                        target.ActualText(
                            drawingGroup.ActualText,
                            (visual.X + drawingGroup.ActualTextAnchorX * scaleX) * PointsPerCssPixel,
                            (visual.Y + drawingGroup.ActualTextAnchorY * scaleY) * PointsPerCssPixel,
                            addGroupPaint);
                    }
                    continue;
                }
                if (element is OfficeDrawingEffectGroup effectGroup) {
                    FlushShapes();
                    if (effectGroup.SoftMask != null) {
                        // The shared PDF drawing writer owns mask composition;
                        // a canvas effect carries only transform and opacity.
                        var masked = new OfficeDrawing(source.Width, source.Height) {
                            TextShapingProvider = source.TextShapingProvider,
                            TextShapingLanguage = source.TextShapingLanguage
                        };
                        masked.Fonts.AddRange(source.Fonts);
                        masked.AddEffectDrawing(effectGroup.Drawing, effectGroup.Transform,
                            effectGroup.BlendMode, effectGroup.SoftMask, effectGroup.Opacity);
                        target.DrawingForClippedRendering(masked, originX, originY,
                            visual.Width * PointsPerCssPixel, visual.Height * PointsPerCssPixel);
                        continue;
                    }
                    OfficeTransform pageTransform = pageToDrawing
                        .Then(effectGroup.Transform)
                        .Then(drawingToPage);
                    OfficeDrawing nestedDrawing = effectGroup.Drawing;
                    target.Effect(pageTransform, effectGroup.Opacity, nested => AddElements(nested, nestedDrawing.Elements));
                    continue;
                }
                if (element is OfficeDrawingTilingPattern tilingPattern) {
                    FlushShapes();
                    OfficeDrawing tileDrawing = tilingPattern.Tile;
                    OfficeImagePlacement area = tilingPattern.Area;
                    double clipX = (visual.X + (area.X * scaleX)) * PointsPerCssPixel;
                    double clipY = (visual.Y + (area.Y * scaleY)) * PointsPerCssPixel;
                    double clipWidth = area.Width * scaleX * PointsPerCssPixel;
                    double clipHeight = area.Height * scaleY * PointsPerCssPixel;
                    target.SourceClip(clipX, clipY, clipWidth, clipHeight, clipped => {
                        foreach (OfficeTransform tileTransform in tilingPattern.GetTileTransforms()) {
                            cancellationToken.ThrowIfCancellationRequested();
                            OfficeTransform pageTransform = pageToDrawing
                                .Then(tileTransform)
                                .Then(drawingToPage);
                            clipped.Effect(pageTransform, tilingPattern.Opacity,
                                nested => AddElements(nested, tileDrawing.Elements));
                        }
                    });
                    continue;
                }
                if (element is OfficeDrawingImage image) {
                    FlushShapes();
                    PdfCore.PdfCanvasImageResource? imageResource = imageResources.GetOrCreate(
                        image.EncodedBytes, image.ContentType);
                    if (imageResource == null) continue;
                    PdfCore.PdfImageStyle? imageStyle = image.Projection.HasCrop
                        ? new PdfCore.PdfImageStyle {
                            SourceCrop = new PdfCore.PdfImageSourceCrop(
                                image.Projection.SourceCrop.Left,
                                image.Projection.SourceCrop.Top,
                                image.Projection.SourceCrop.Right,
                                image.Projection.SourceCrop.Bottom)
                        }
                        : null;
                    Action<PdfCore.PdfPageCanvas> addImage = imageTarget => imageTarget.ImageShared(
                        imageResource,
                        (visual.X + image.Projection.X * scaleX) * PointsPerCssPixel,
                        (visual.Y + image.Projection.Y * scaleY) * PointsPerCssPixel,
                        image.Projection.Width * scaleX * PointsPerCssPixel,
                        image.Projection.Height * scaleY * PointsPerCssPixel,
                        imageStyle,
                        alternativeText: string.IsNullOrWhiteSpace(image.AlternativeText) ? null : image.AlternativeText);
                    OfficeTransform imageTransform = image.Projection.CreateFrameTransform().CreateDestinationTransform();
                    if (imageTransform == OfficeTransform.Identity && image.Opacity >= 1D) {
                        addImage(target);
                    } else {
                        OfficeTransform pageImageTransform = pageToDrawing.Then(imageTransform).Then(drawingToPage);
                        target.Effect(pageImageTransform, image.Opacity, addImage);
                    }
                    continue;
                }
                if (element is OfficeDrawingLink link) {
                    FlushShapes();
                    double linkX = (visual.X + link.X * scaleX) * PointsPerCssPixel;
                    double linkY = (visual.Y + link.Y * scaleY) * PointsPerCssPixel;
                    double linkWidth = link.Width * scaleX * PointsPerCssPixel;
                    double linkHeight = link.Height * scaleY * PointsPerCssPixel;
                    if (IsFragmentLink(link.Uri)) {
                        target.LinkToNamedDestination(
                            MapNamedDestination(DecodeFragmentName(link.Uri.Substring(1))),
                            linkX,
                            linkY,
                            linkWidth,
                            linkHeight,
                            link.AlternativeText);
                    } else {
                        OfficeShape linkArea = OfficeShape.Rectangle(linkWidth, linkHeight);
                        linkArea.FillColor = null;
                        linkArea.StrokeColor = null;
                        target.Shape(linkArea, linkX, linkY, style: new PdfCore.PdfDrawingStyle { Decorative = true }, linkUri: link.Uri, linkContents: link.AlternativeText);
                    }
                    continue;
                }
                if (element is not OfficeDrawingText text || text.Text.Length == 0) continue;
                FlushShapes();
                double textY = visual.Y + text.Y * scaleY;
                var projectedText = new HtmlRenderText(
                    text.RasterText,
                    visual.X + text.X * scaleX,
                    textY,
                    text.Width * scaleX,
                    text.Height * scaleY,
                    text.Font.WithSize(text.Font.Size * scaleY),
                    text.Color ?? OfficeColor.Black,
                    text.Alignment,
                    (text.LineHeight ?? text.Font.Size * 1.2D) * scaleY,
                    paintOrder: 0,
                    linkUri: null,
                    source: visual.Source,
                    semanticRole: "span",
                    layoutY: textY,
                    semanticNodeId: null,
                    textAdvanceWidth: text.TextAdvanceWidth * scaleX,
                    underlineStyle: text.UnderlineStyle,
                    strikethroughStyle: text.StrikethroughStyle,
                    baseline: text.Baseline,
                    baselineLevel: text.BaselineLevel,
                    baselineScale: text.BaselineScale,
                    // Drawing baseline offsets move ink within its frame; they must not
                    // translate the frame/clip as a CSS script displacement would.
                    baselineOffset: 0D,
                    decorationColor: text.DecorationColor,
                    featureSettings: text.FeatureSettings,
                    fontPalette: text.FontPalette);
                void AddDrawingText(PdfCore.PdfPageCanvas textTarget) => AddText(textTarget, projectedText, webFonts, conversionReport,
                    visual.X + visual.Width, asSpan: true, logicalTextOwned: text.Text != text.RasterText, cancellationToken,
                    baselineFontSize: Math.Max(0.001D, (text.Font.Size + text.BaselineOffset) * scaleY),
                    preservePositionedFrame: true);
                if (text.Text == text.RasterText) AddDrawingText(target);
                else target.ActualText(text.Text, AddDrawingText);
            }
            FlushShapes();
        }

        if (source.Elements.Count == 0) return;
        if (string.IsNullOrWhiteSpace(visual.AlternativeText)) {
            AddElements(canvas, source.Elements);
        } else {
            canvas.Figure(visual.AlternativeText!, figure => AddElements(figure, source.Elements));
        }
        double linkX = visual.ImageX * PointsPerCssPixel;
        double linkY = visual.ImageY * PointsPerCssPixel;
        double linkWidth = visual.ImageWidth * PointsPerCssPixel;
        double linkHeight = visual.ImageHeight * PointsPerCssPixel;
        if (fragmentLink) {
            canvas.LinkToNamedDestination(
                MapNamedDestination(visual.LinkUri!.Substring(1)),
                linkX,
                linkY,
                linkWidth,
                linkHeight,
                visual.Source);
        } else if (visual.LinkUri != null && !suppressLink) {
            // The HTML anchor covers the visible drawing viewport, including whitespace
            // between SVG shapes. Nested groups may clip their own paint and annotations.
            OfficeShape linkArea = OfficeShape.Rectangle(linkWidth, linkHeight);
            linkArea.FillColor = null;
            linkArea.StrokeColor = null;
            canvas.Shape(linkArea, linkX, linkY, style: new PdfCore.PdfDrawingStyle { Decorative = true }, linkUri: visual.LinkUri, linkContents: visual.Source);
        }
    }

}

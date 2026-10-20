using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private bool TryDrawSvgImageAt(OfficeDrawingImage image, double originX, double originTopY, long? logicalOrder = null) {
            bool isSvg = OfficeImageInfo.FromMimeType(image.ContentType) == OfficeImageFormat.Svg ||
                (OfficeImageReader.TryIdentifyByContent(image.EncodedBytes, null, out OfficeImageInfo info) && info.Format == OfficeImageFormat.Svg);
            if (!isSvg) return false;
            OfficeDrawing vector = ReadSvgImageForPdf(image.EncodedBytes);
            if (!CanDrawSvgSceneAsVector(vector)) {
                currentOpts.AddLayoutDiagnostic("SvgImageRasterFallback", "Drawing image",
                    "SVG image paint not supported by the PDF vector writer was rasterized.", PdfLayoutDiagnosticKind.SimplifiedContent);
                byte[] png = RasterizeSvgImageForPdf(vector, image.Projection);
                DrawDrawingImageAt(OfficeDrawingImage.CreateWithInterpolation(png, "image/png", image.Projection, image.Interpolate,
                    image.AlternativeText, image.Opacity), originX, originTopY, logicalOrder);
                return true;
            }

            OfficeImageProjection projection = image.Projection;
            bool markedContent;
            AppendDrawingMarkedContentBegin(new PdfDrawingStyle { AlternativeText = image.AlternativeText }, out markedContent, logicalOrder: logicalOrder);
            bool previousWrappers = _suppressCanvasAccessibilityWrappers;
            bool previousLinks = _suppressDrawingLinks;
            _suppressCanvasAccessibilityWrappers = true;
            _suppressDrawingLinks = true;
            try {
                RenderEffectGroup(ToTopLeftPageTransform(projection.CreateFrameTransform().CreateDestinationTransform(), originX, originTopY), image.Opacity, () => {
                    var content = new ContentStreamBuilder(sb).SaveState();
                    AppendClipPath(sb, OfficeClipPath.Rectangle(projection.Width, projection.Height),
                        originX + projection.X, originTopY - projection.Y - projection.Height, projection.Height);
                    double width = projection.Width / projection.SourceWidth;
                    double height = projection.Height / projection.SourceHeight;
                    OfficeTransform placement = OfficeTransform.Scale(width / vector.Width, height / vector.Height)
                        .Then(OfficeTransform.Translate(projection.X - projection.SourceLeft * width, projection.Y - projection.SourceTop * height));
                    RenderEffectGroup(ToTopLeftPageTransform(placement, originX, originTopY), 1D,
                        () => DrawDrawingElements(vector, originX, originTopY));
                    content.RestoreState();
                });
            } finally {
                _suppressCanvasAccessibilityWrappers = previousWrappers;
                _suppressDrawingLinks = previousLinks;
                AppendDrawingMarkedContentEnd(markedContent);
            }
            return true;
        }

        private OfficeDrawing ReadSvgImageForPdf(byte[] bytes) {
            return PdfSvgImageReader.Read(bytes, cancellationToken);
        }

        private byte[] RasterizeSvgImageForPdf(OfficeDrawing vector, OfficeImageProjection projection) {
            double desiredScale = Math.Max(1D, Math.Max(projection.Width / vector.Width, projection.Height / vector.Height));
            OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(vector.Width, vector.Height, OfficeImageExportFormat.Png,
                new OfficeImageExportOptions { Scale = desiredScale, MaximumRasterPixels = 16_000_000L,
                    RasterOverflowBehavior = OfficeRasterOverflowBehavior.ReduceScale });
            return OfficeDrawingRasterRenderer.ToPng(vector, new OfficeDrawingRasterRenderOptions {
                Scale = plan.Limit.Scale, Background = OfficeColor.Transparent,
                CancellationToken = cancellationToken, MaximumRasterPixels = 16_000_000L,
                ThrowOnImageDecodeFailure = true
            });
        }

        private static bool CanDrawSvgSceneAsVector(OfficeDrawing drawing) {
            foreach (OfficeDrawingElement element in drawing.Elements) {
                if (element is OfficeDrawingGroup group) {
                    if (!CanDrawSvgSceneAsVector(group.InnerDrawing)) return false;
                } else if (element is OfficeDrawingEffectGroup effect) {
                    if (effect.SoftMask is { } mask && (mask.Mode != OfficeSoftMaskMode.Alpha || mask.BackdropColor.A != 0 || !CanDrawSvgSceneAsVector(mask.InnerDrawing))) return false;
                    if (!CanDrawSvgSceneAsVector(effect.InnerDrawing)) return false;
                } else if (element is OfficeDrawingTilingPattern pattern) {
                    if (!CanDrawSvgSceneAsVector(pattern.InnerTile)) return false;
                } else if (element is not (OfficeDrawingShape or OfficeDrawingImage or OfficeDrawingText or OfficeDrawingRichText or OfficeDrawingLink)) return false;
            }
            return true;
        }
    }
}

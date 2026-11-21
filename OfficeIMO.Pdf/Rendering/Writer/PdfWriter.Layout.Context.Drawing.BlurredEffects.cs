using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        // A blurred edge needs no more detail than a 150 DPI print shows.
        private const double BlurredEffectPixelsPerPoint = 150D / 72D;

        private readonly Dictionary<(byte[] Coverage, global::OfficeIMO.OfficeColor Color), PdfDocument.PreparedImage> _blurredEffectImages = new();

        /// <summary>
        /// Paints a soft shadow or glow as one image of the blurred silhouette: the effect color with the blur
        /// as its soft mask, placed in the shape's local space so transforms and offsets match the vector shape.
        /// </summary>
        private void DrawBlurredShapeEffectAt(OfficeBlurredShapeEffect effect, double xShape, double bottomY, double shapeHeight) {
            cancellationToken.ThrowIfCancellationRequested();
            double transformScale = OfficeBlurredShapeEffect.GetAxisScale(effect.Transform);
            if (!(transformScale > 0D)) return;
            double density = effect.ChoosePixelsPerUnit(BlurredEffectPixelsPerPoint * transformScale,
                OfficeBlurredShapeEffect.MaximumRasterPixels);
            // Unspecified caps and joins keep the PDF defaults the vector shape uses.
            byte[] coverage = effect.RenderCoverage(density, cancellationToken, out int pixelWidth, out int pixelHeight,
                OfficeStrokeLineCap.Butt, OfficeStrokeLineJoin.Miter);
            // Equal primitive effects share one coverage array, so a repeated theme effect is encoded once per document.
            var cacheKey = (coverage, effect.Color);
            if (!_blurredEffectImages.TryGetValue(cacheKey, out PdfDocument.PreparedImage prepared)) {
                var rgba = new byte[coverage.Length * 4];
                for (int i = 0, offset = 0; i < coverage.Length; i++, offset += 4) {
                    rgba[offset] = effect.Color.R; rgba[offset + 1] = effect.Color.G; rgba[offset + 2] = effect.Color.B; rgba[offset + 3] = coverage[i];
                }
                // A real PNG keeps image optimization and PDF/X color conversion working on the effect like on any
                // other image; preparing it splits the blur into a soft mask under a flat color image.
                byte[] png = OfficePngWriter.EncodeRgba(pixelWidth, pixelHeight, rgba, OfficePngCompression.Fastest);
                prepared = PdfDocument.PrepareImageBytes(png, cancellationToken);
                _blurredEffectImages[cacheKey] = prepared;
            }
            var imageStyle = new PdfImageStyle { Fit = OfficeImageFit.Stretch };
            var block = new ImageBlock(prepared.Data, 1D, 1D, prepared.Info, imageStyle,
                useDataSnapshot: true, preparedStream: prepared.PreparedStream);
            // The image is placed on the unit square; the matrices below map that square to the effect's local box.
            PageImage pageImage = CreatePageImage(block, imageStyle, 0D, 0D, 1D, 1D);
            pageImage.Interpolate = true;
            pageImage.IsOpacityData = _drawingIsOpacityData;
            pageImage.Opacity = effect.Opacity;
            pageImage.GraphicsStateName = EnsureGraphicsState(effect.Opacity, effect.Opacity);
            pageImage.SuppressAccessibilityWrapper = true;

            double width = pixelWidth / density, height = pixelHeight / density;
            OfficeTransform localTransform = effect.Transform ?? OfficeTransform.Identity;
            double originX = xShape + effect.OffsetX, originY = bottomY - effect.OffsetY;
            // Page footprint of the placed image, so canvas clipping sees where the effect really paints.
            double left = double.PositiveInfinity, bottom = double.PositiveInfinity, right = double.NegativeInfinity, top = double.NegativeInfinity;
            foreach ((double u, double v) in new[] {
                (effect.LocalLeft, effect.LocalTop), (effect.LocalLeft + width, effect.LocalTop),
                (effect.LocalLeft, effect.LocalTop + height), (effect.LocalLeft + width, effect.LocalTop + height) }) {
                double x = originX + localTransform.OffsetX + localTransform.M11 * u + localTransform.M21 * v;
                double y = originY + shapeHeight - localTransform.OffsetY - localTransform.M12 * u - localTransform.M22 * v;
                left = Math.Min(left, x); right = Math.Max(right, x);
                bottom = Math.Min(bottom, y); top = Math.Max(top, y);
            }
            pageImage.EffectiveX = left;
            pageImage.EffectiveY = bottom;
            pageImage.EffectiveW = right - left;
            pageImage.EffectiveH = top - bottom;
            currentPage!.Images.Add(pageImage);
            pageImage.InlineDrawToken = AllocateInlineImageDrawToken(currentPage);

            var content = new ContentStreamBuilder(sb).SaveState();
            // Shape-local space with y down, offset like the shadow's vector geometry.
            ApplyLocalTransform(sb, localTransform, originX, originY, shapeHeight);
            if (effect.ClipPath != null) AppendLocalClipPath(sb, effect.ClipPath);
            content.TransformMatrix(width, 0D, 0D, -height, effect.LocalLeft, effect.LocalTop + height);
            sb.Append(pageImage.InlineDrawToken);
            content.RestoreState();
            pageDirty = true;
        }
    }
}

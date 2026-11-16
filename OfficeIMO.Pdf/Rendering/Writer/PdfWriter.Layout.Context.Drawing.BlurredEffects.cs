using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        // A blurred edge needs no more detail than a 150 DPI print shows.
        private const double BlurredEffectPixelsPerPoint = 150D / 72D;

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
            byte[] coverage = effect.RenderCoverage(density, cancellationToken, out int pixelWidth, out int pixelHeight);
            PdfImageStream stream = CreateBlurredEffectImage(effect.Color, coverage, pixelWidth, pixelHeight, cancellationToken);
            var imageStyle = new PdfImageStyle { Fit = OfficeImageFit.Stretch };
            var block = new ImageBlock(stream.SoftMask!.Data, 1D, 1D, new OfficeImageInfo(OfficeImageFormat.Png, pixelWidth, pixelHeight),
                imageStyle, useDataSnapshot: false, preparedStream: stream);
            // The image is placed on the unit square; the matrices below map that square to the effect's local box.
            PageImage pageImage = CreatePageImage(block, imageStyle, 0D, 0D, 1D, 1D);
            pageImage.Interpolate = true;
            pageImage.IsOpacityData = _drawingIsOpacityData;
            pageImage.Opacity = effect.Opacity;
            pageImage.GraphicsStateName = EnsureGraphicsState(effect.Opacity, effect.Opacity);
            pageImage.SuppressAccessibilityWrapper = true;
            currentPage!.Images.Add(pageImage);
            pageImage.InlineDrawToken = AllocateInlineImageDrawToken(currentPage);

            double width = pixelWidth / density, height = pixelHeight / density;
            var content = new ContentStreamBuilder(sb).SaveState();
            // Shape-local space with y down, offset like the shadow's vector geometry.
            ApplyLocalTransform(sb, effect.Transform ?? OfficeTransform.Identity,
                xShape + effect.OffsetX, bottomY - effect.OffsetY, shapeHeight);
            if (effect.ClipPath != null) AppendLocalClipPath(sb, effect.ClipPath);
            content.TransformMatrix(width, 0D, 0D, -height, effect.LocalLeft, effect.LocalTop + height);
            sb.Append(pageImage.InlineDrawToken);
            content.RestoreState();
            pageDirty = true;
        }

        // A flat color image with the blurred coverage as its soft mask. Mask rows use the PNG Up predictor,
        // which leaves little but zeros on a smooth blur.
        private static PdfImageStream CreateBlurredEffectImage(global::OfficeIMO.OfficeColor color, byte[] coverage,
            int width, int height, CancellationToken cancellationToken) {
            int colorRowLength = width * 3 + 1, maskRowLength = width + 1;
            var colorRows = new byte[colorRowLength * height];
            for (int x = 0; x < width; x++) {
                colorRows[1 + x * 3] = color.R; colorRows[2 + x * 3] = color.G; colorRows[3 + x * 3] = color.B;
            }
            for (int y = 1; y < height; y++) colorRows[y * colorRowLength] = 2;
            var maskRows = new byte[maskRowLength * height];
            Buffer.BlockCopy(coverage, 0, maskRows, 1, width);
            for (int y = 1; y < height; y++) {
                cancellationToken.ThrowIfCancellationRequested();
                int row = y * maskRowLength, source = y * width;
                maskRows[row] = 2;
                for (int x = 0; x < width; x++) maskRows[row + 1 + x] = unchecked((byte)(coverage[source + x] - coverage[source - width + x]));
            }
            return new PdfImageStream {
                Data = DeflateZlib(colorRows, cancellationToken),
                PixelWidth = width,
                PixelHeight = height,
                DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceRGB", 3, width),
                SoftMask = new PdfImageStream {
                    Data = DeflateZlib(maskRows, cancellationToken),
                    PixelWidth = width,
                    PixelHeight = height,
                    DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceGray", 1, width)
                }
            };
        }
    }
}

using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static bool TryConvertImageStreamToCmyk(
        byte[] sourceBytes,
        PdfImageStream image,
        PdfPrintColorTransform transform,
        bool flattenTransparency,
        PdfColor transparencyBackground,
        out string? unsupportedReason) {
        unsupportedReason = null;
        long declaredPixelCount = (long)image.PixelWidth * image.PixelHeight;
        if (declaredPixelCount <= 0 || declaredPixelCount > MaxPngPixelCount) {
            unsupportedReason = "Raster image dimensions exceed the bounded PDF/X CMYK conversion limit.";
            return false;
        }

        if (!OfficeRasterImageDecoder.TryDecode(sourceBytes, out OfficeRasterImage? raster) || raster == null) {
            unsupportedReason = "The raster image decoder could not produce RGBA pixels for CMYK conversion.";
            return false;
        }

        long pixelCount = (long)raster.Width * raster.Height;
        if (pixelCount <= 0 || pixelCount > MaxPngPixelCount) {
            unsupportedReason = "Raster image dimensions exceed the bounded PDF/X CMYK conversion limit.";
            return false;
        }

        byte[] pixels = raster.GetPixels();
        int baseRowLength = checked(1 + raster.Width * 4);
        byte[] cmykRows = new byte[checked(baseRowLength * raster.Height)];
        byte[]? alphaRows = flattenTransparency ? null : new byte[checked((1 + raster.Width) * raster.Height)];
        PdfColor background = transparencyBackground;
        var components = new double[4];
        bool hasTransparency = false;

        for (int row = 0; row < raster.Height; row++) {
            int sourceRow = row * raster.Width * 4;
            int targetRow = row * baseRowLength;
            int alphaRow = row * (1 + raster.Width);
            cmykRows[targetRow] = 0;
            if (alphaRows != null) alphaRows[alphaRow] = 0;
            for (int column = 0; column < raster.Width; column++) {
                int source = sourceRow + column * 4;
                byte alpha = pixels[source + 3];
                hasTransparency |= alpha != 255;
                OfficeColor color;
                if (flattenTransparency && alpha != 255) {
                    double opacity = alpha / 255D;
                    color = OfficeColor.FromRgb(
                        Composite(pixels[source], background.R, opacity),
                        Composite(pixels[source + 1], background.G, opacity),
                        Composite(pixels[source + 2], background.B, opacity));
                } else {
                    color = OfficeColor.FromRgb(pixels[source], pixels[source + 1], pixels[source + 2]);
                }

                transform.Convert(color, components);
                int target = targetRow + 1 + column * 4;
                cmykRows[target] = ToComponentByte(components[0]);
                cmykRows[target + 1] = ToComponentByte(components[1]);
                cmykRows[target + 2] = ToComponentByte(components[2]);
                cmykRows[target + 3] = ToComponentByte(components[3]);
                if (alphaRows != null) alphaRows[alphaRow + 1 + column] = alpha;
            }
        }

        image.Data = DeflateZlib(cmykRows);
        image.PixelWidth = raster.Width;
        image.PixelHeight = raster.Height;
        image.DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceCMYK", 4, raster.Width);
        image.SoftMask = hasTransparency && alphaRows != null
            ? new PdfImageStream {
                Data = DeflateZlib(alphaRows),
                PixelWidth = raster.Width,
                PixelHeight = raster.Height,
                DictionarySuffix = BuildPngPredictorDictionarySuffix("/DeviceGray", 1, raster.Width)
            }
            : null;
        return true;
    }

    private static byte Composite(byte foreground, double background, double opacity) =>
        (byte)Math.Round((foreground * opacity) + (background * 255D * (1D - opacity)));

    private static byte ToComponentByte(double value) =>
        (byte)Math.Round(Math.Max(0D, Math.Min(1D, value)) * 255D);
}

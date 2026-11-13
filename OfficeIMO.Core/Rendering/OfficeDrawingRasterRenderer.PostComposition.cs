namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static void ApplyPostCompositionColorTransform(OfficeRasterImage image,
        IOfficePostCompositionColorTransform? transform, CancellationToken cancellationToken) {
        if (transform == null) return;
        byte[] pixels = image.PixelBuffer;
        int stride = checked(image.Width * 4);
        for (int y = 0; y < image.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            int rowEnd = checked((y + 1) * stride);
            for (int offset = y * stride; offset < rowEnd; offset += 4) {
                if ((offset & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                byte alpha = pixels[offset + 3];
                if (alpha == 0) continue;
                OfficeColor source = OfficeColor.FromRgba(pixels[offset], pixels[offset + 1], pixels[offset + 2], alpha);
                if (!transform.TryTransform(source, out OfficeColor converted)) continue;
                pixels[offset] = converted.R;
                pixels[offset + 1] = converted.G;
                pixels[offset + 2] = converted.B;
                // A color transform cannot change surface coverage.
            }
        }
    }
}

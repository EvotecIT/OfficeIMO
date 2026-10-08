using OfficeIMO.Drawing;
using System.IO;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private static void AddFunctionShading(OfficeDrawing drawing, PdfPageVisualPrimitive primitive,
        PdfTextClippingBudget clippingBudget) {
        if (primitive.Width <= 0D || primitive.Height <= 0D) return;
        var clip = PdfPageClipPath.Rectangle(primitive.X, primitive.Y, primitive.Width, primitive.Height);
        if (primitive.ClipPath.HasValue) clip = clippingBudget.ResolveActiveClip(primitive.ClipPath, clip);
        if (!TryFitClipToDrawing(clip, drawing.Width, drawing.Height, out var fitted)) return;
        var path = fitted.ToOfficeClipPath(fitted.X, fitted.Y);
        if (path == null) return;
        var paint = primitive.FunctionPaint!;
        var field = paint.Resource;
        double scale = field.RasterScale;
        if (!(scale > 0D) || double.IsInfinity(scale)) throw new InvalidOperationException("Function shading requires a positive finite raster scale.");
        // Align sampling with the output pixel grid even for clipped partial pixels.
        double leftPixel = Math.Floor(fitted.X * scale), topPixel = Math.Floor(fitted.Y * scale);
        double width = Math.Ceiling((fitted.X + fitted.Width) * scale) - leftPixel;
        double height = Math.Ceiling((fitted.Y + fitted.Height) * scale) - topPixel;
        if (!IsFinite(width) || !IsFinite(height) || width > int.MaxValue || height > int.MaxValue ||
            width * height > int.MaxValue / 4D)
            throw PdfReadLimitException.Create(PdfReadLimitKind.FunctionShadingPixels, int.MaxValue / 4, long.MaxValue);
        if (width <= 0D || height <= 0D) return;
        int pixelWidth = (int)width, pixelHeight = (int)height;
        field.ChargePixels?.Invoke((long)pixelWidth * pixelHeight);
        var raster = new OfficeRasterImage(pixelWidth, pixelHeight);
        var input = new double[2]; var components = new double[field.ComponentCount];
        for (int y = 0; y < pixelHeight; y++) for (int x = 0; x < pixelWidth; x++) {
            var point = paint.InverseTransform.TransformPoint(new OfficePoint(
                (leftPixel + x + .5D) / scale, (topPixel + y + .5D) / scale));
            if (!field.TrySample(point.X, point.Y, input, components, out var color))
                throw new InvalidDataException("PDF function shading could not evaluate a point in its domain.");
            raster.SetPixel(x, y, color);
        }
        byte[] png = OfficeRasterImageEncoder.Encode(raster, OfficeImageExportFormat.Png, null, 256L * 1024 * 1024, field.CancellationToken);
        drawing.AddClippedImageSharedWithInterpolation(png, "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(leftPixel / scale, topPixel / scale, width / scale, height / scale)),
            interpolate: false, fitted.X, fitted.Y, path, opacity: primitive.FillOpacity ?? 1D);
    }
}

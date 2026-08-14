using System;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgImageRenderer {
    private const int MaximumVectorizedNearestNeighborRectangles = 1_000_000;

    private static StringBuilder AppendNearestNeighborRaster(
        StringBuilder builder,
        OfficeRasterImage? raster,
        SvgImageLayout layout,
        string? clipPathId,
        System.Threading.CancellationToken cancellationToken) {
        if (raster == null) {
            throw new InvalidOperationException("SVG export cannot preserve nearest-neighbor sampling for an undecodable image.");
        }

        if (!string.IsNullOrEmpty(clipPathId) && layout.EffectiveClip != null) {
            OfficeImagePlacement clip = layout.EffectiveClip.Value;
            builder.AppendRectClipPathDefinition(clipPathId!, clip.X, clip.Y, clip.Width, clip.Height);
        }

        builder.Append("<g");
        if (!string.IsNullOrEmpty(clipPathId) && layout.EffectiveClip != null) {
            builder.AppendClipPathReference(clipPathId!);
        }
        if (layout.Transform != null) {
            builder.AppendAttribute("transform", layout.Transform);
        }
        builder.Append("><g shape-rendering=\"crispEdges\" transform=\"translate(")
            .Append(OfficeSvgFormatting.FormatNumber(layout.ImagePlacement.X)).Append(' ')
            .Append(OfficeSvgFormatting.FormatNumber(layout.ImagePlacement.Y)).Append(") scale(")
            .Append(OfficeSvgFormatting.FormatNumber(layout.ImagePlacement.Width / raster.Width)).Append(' ')
            .Append(OfficeSvgFormatting.FormatNumber(layout.ImagePlacement.Height / raster.Height)).Append(")\">");

        byte[] pixels = raster.GetPixels();
        GetVisibleSourceBounds(raster, layout, out int minimumX, out int minimumY, out int maximumX, out int maximumY);
        int rectangleCount = 0;
        for (int y = minimumY; y < maximumY; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            int x = minimumX;
            while (x < maximumX) {
                int offset = ((y * raster.Width) + x) * 4;
                byte red = pixels[offset];
                byte green = pixels[offset + 1];
                byte blue = pixels[offset + 2];
                byte alpha = pixels[offset + 3];
                int end = x + 1;
                while (end < maximumX) {
                    int next = ((y * raster.Width) + end) * 4;
                    if (pixels[next] != red || pixels[next + 1] != green || pixels[next + 2] != blue || pixels[next + 3] != alpha) break;
                    end++;
                }
                if (alpha != 0) {
                    if (++rectangleCount > MaximumVectorizedNearestNeighborRectangles) {
                        throw new InvalidOperationException("SVG nearest-neighbor image export exceeds the supported vectorization limit.");
                    }
                    var color = OfficeColor.FromRgba(red, green, blue, alpha);
                    builder.Append("<rect x=\"").Append(x).Append("\" y=\"").Append(y)
                        .Append("\" width=\"").Append(end - x).Append("\" height=\"1\"")
                        .AppendPaintAttribute("fill", color)
                        .Append("/>");
                }
                x = end;
            }
        }

        return builder.Append("</g></g>");
    }

    private static void GetVisibleSourceBounds(
        OfficeRasterImage raster,
        SvgImageLayout layout,
        out int minimumX,
        out int minimumY,
        out int maximumX,
        out int maximumY) {
        minimumX = 0;
        minimumY = 0;
        maximumX = raster.Width;
        maximumY = raster.Height;
        if (layout.EffectiveClip is not OfficeImagePlacement clip) return;

        OfficeImagePlacement image = layout.ImagePlacement;
        double visibleLeft = Math.Max(image.X, clip.X);
        double visibleTop = Math.Max(image.Y, clip.Y);
        double visibleRight = Math.Min(image.X + image.Width, clip.X + clip.Width);
        double visibleBottom = Math.Min(image.Y + image.Height, clip.Y + clip.Height);
        if (visibleRight <= visibleLeft || visibleBottom <= visibleTop) {
            maximumX = minimumX;
            maximumY = minimumY;
            return;
        }

        minimumX = ClampSourceIndex((int)Math.Floor((visibleLeft - image.X) * raster.Width / image.Width), raster.Width);
        minimumY = ClampSourceIndex((int)Math.Floor((visibleTop - image.Y) * raster.Height / image.Height), raster.Height);
        maximumX = ClampSourceIndex((int)Math.Ceiling((visibleRight - image.X) * raster.Width / image.Width), raster.Width);
        maximumY = ClampSourceIndex((int)Math.Ceiling((visibleBottom - image.Y) * raster.Height / image.Height), raster.Height);
    }

    private static int ClampSourceIndex(int value, int length) => Math.Max(0, Math.Min(length, value));

}

using System;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgImageRenderer {
    private const int MaximumVectorizedNearestNeighborPixels = 1_000_000;

    private static StringBuilder AppendNearestNeighborRaster(
        StringBuilder builder,
        OfficeRasterImage? raster,
        SvgImageLayout layout,
        string? clipPathId) {
        if (raster == null) {
            throw new InvalidOperationException("SVG export cannot preserve nearest-neighbor sampling for an undecodable image.");
        }
        long destinationPixels = (long)Math.Ceiling(layout.ImagePlacement.Width) *
            (long)Math.Ceiling(layout.ImagePlacement.Height);
        if ((long)raster.Width * raster.Height > MaximumVectorizedNearestNeighborPixels &&
            destinationPixels > 0 &&
            destinationPixels < (long)raster.Width * raster.Height) {
            int width = checked((int)Math.Ceiling(layout.ImagePlacement.Width));
            int height = checked((int)Math.Ceiling(layout.ImagePlacement.Height));
            if (width <= 0 || height <= 0 || (long)width * height > MaximumVectorizedNearestNeighborPixels) {
                throw new InvalidOperationException("SVG nearest-neighbor image export exceeds the supported vectorization limit.");
            }
            raster = OfficeRasterResampler.Resize(
                raster,
                width,
                height,
                OfficeRasterResamplingMode.NearestNeighbor);
        }
        if ((long)raster.Width * raster.Height > MaximumVectorizedNearestNeighborPixels) {
            throw new InvalidOperationException("SVG nearest-neighbor image export exceeds the supported vectorization limit.");
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
        for (int y = 0; y < raster.Height; y++) {
            int x = 0;
            while (x < raster.Width) {
                int offset = ((y * raster.Width) + x) * 4;
                byte red = pixels[offset];
                byte green = pixels[offset + 1];
                byte blue = pixels[offset + 2];
                byte alpha = pixels[offset + 3];
                int end = x + 1;
                while (end < raster.Width) {
                    int next = ((y * raster.Width) + end) * 4;
                    if (pixels[next] != red || pixels[next + 1] != green || pixels[next + 2] != blue || pixels[next + 3] != alpha) break;
                    end++;
                }
                if (alpha != 0) {
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

}

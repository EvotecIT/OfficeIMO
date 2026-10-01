using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryApplySvgFilterGraph(OfficeDrawing source, SvgFilterGraph graph,
        SvgElementReferenceRegistry references, OfficeTransform transform, out OfficeDrawing result) {
        result = source;
        // A raster filter must never silently discard searchable text or interactive metadata.
        // Rotation/shear needs a filter-local sampling grid, rather than axis-aligned blur kernels.
        if (ContainsSvgFilterSemantics(source) || Math.Abs(transform.M12) > 0.000001D ||
            Math.Abs(transform.M21) > 0.000001D || !transform.TryInvert(out OfficeTransform inverse) ||
            !TryGetSvgFilterGeometryBounds(source.Elements, out SvgInteractiveBounds bounds) ||
            !TryResolveFilterRegion(graph, source, bounds.Transform(inverse), transform, out SvgInteractiveBounds region)) return false;

        if (region.Right <= region.Left || region.Bottom <= region.Top) {
            result = CreateEmptySvgFilterResult(source, bounds);
            return true;
        }
        double left = Math.Floor(region.Left), top = Math.Floor(region.Top);
        double w = Math.Ceiling(region.Right) - left, h = Math.Ceiling(region.Bottom) - top;
        if (w <= 0D || h <= 0D) {
            result = CreateEmptySvgFilterResult(source, bounds);
            return true;
        }
        if (!FiniteFilterNumber(w) || !FiniteFilterNumber(h) || w > int.MaxValue || h > int.MaxValue) return false;
        double pixels = w * h;
        double work = pixels * (4D + EstimateSvgFilterSourceWork(source, references.CancellationToken));
        foreach (SvgFilterNode node in graph.Nodes) {
            if (node.Operation == SvgFilterOperation.Blur) {
                double sx = node.X * Math.Abs(transform.M11), sy = node.Y * Math.Abs(transform.M22);
                if (sx > 64D || sy > 64D) return false;
                work += pixels * (FilterKernelLength(sx) + FilterKernelLength(sy));
            } else {
                work += pixels * (node.Operation == SvgFilterOperation.Matrix ? 20D : 8D);
            }
        }
        // RGBA float results cost four ordinary surfaces each. Include source/alpha,
        // separable temporary, final RGBA, encoder and retained PNG copies before allocating.
        if (!references.TryChargeFilterGraph(pixels, graph.Nodes.Count * 4 + 24, work)) return false;
        int width = (int)w, height = (int)h;
        System.Threading.CancellationToken token = references.CancellationToken;
        token.ThrowIfCancellationRequested();
        try {
            // Render the complete filter region, even outside the SVG viewport. Cropping
            // the input at the viewport first would lose paint later moved into view.
            var canvas = new OfficeDrawing(w, h);
            canvas.AddDrawingForClippedRendering(source, -left, -top, null);
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(canvas, new OfficeDrawingRasterRenderOptions {
                MaximumRasterPixels = (long)SvgElementReferenceRegistry.MaximumIntermediateSurfacePixels,
                ThrowOnImageDecodeFailure = true,
                CancellationToken = token
            });
            float[] graphic = ReadFilterPixels(raster, graph.Linear, token);
            ClipFilterPixels(graphic, width, height, region, left, top, token);
            float[] alpha = new float[graphic.Length];
            for (int y = 0; y < height; y++) {
                token.ThrowIfCancellationRequested();
                for (int i = y * width * 4 + 3; i < (y + 1) * width * 4; i += 4) alpha[i] = graphic[i];
            }
            var outputs = new List<float[]>(graph.Nodes.Count);
            foreach (SvgFilterNode node in graph.Nodes) {
                token.ThrowIfCancellationRequested();
                float[] input = ResolveFilterPixels(node.Input, graphic, alpha, outputs);
                float[] output;
                switch (node.Operation) {
                    case SvgFilterOperation.Blur:
                        output = BlurFilterPixels(input, width, height, node.X * Math.Abs(transform.M11),
                            node.Y * Math.Abs(transform.M22), token);
                        break;
                    case SvgFilterOperation.Offset:
                        output = OffsetFilterPixels(input, width, height, node.X * transform.M11, node.Y * transform.M22, token);
                        break;
                    case SvgFilterOperation.Matrix:
                        output = MatrixFilterPixels(input, width, height, node.Matrix!, token);
                        break;
                    default:
                        output = CompositeFilterPixels(input, ResolveFilterPixels(node.Input2, graphic, alpha, outputs),
                            width, height, node.Operation == SvgFilterOperation.Blend ? node.BlendMode : OfficeBlendMode.Normal, token);
                        break;
                }
                ClipFilterPixels(output, width, height, region, left, top, token);
                outputs.Add(output);
            }
            OfficeRasterImage painted = WriteFilterPixels(outputs[outputs.Count - 1], width, height, graph.Linear, token);
            byte[] png = OfficeRasterImageEncoder.Encode(painted, OfficeImageExportFormat.Png, null, 16L * 1024L * 1024L, token);
            var tile = new OfficeDrawing(w, h);
            tile.AddImage(png, "image/png", new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, w, h)));
            var filtered = new OfficeDrawing(source.Width, source.Height);
            filtered.AddClippedDrawingForRendering(tile, region.Left, region.Top,
                OfficeClipPath.Rectangle(region.Right - region.Left, region.Bottom - region.Top),
                left - region.Left, top - region.Top);
            ((OfficeDrawingGroup)filtered.Elements[0]).UnfilteredGeometryBounds =
                (bounds.Left - region.Left, bounds.Top - region.Top, bounds.Right - region.Left, bounds.Bottom - region.Top);
            result = filtered;
            return true;
        } catch (InvalidOperationException) {
            // Existing managed renderer/encoder limits or unsupported source paint.
            // The importer reports the unapplied filter and retains its source drawing.
            return false;
        } catch (NotSupportedException) {
            // Recognized image formats may require an optional codec. Graph import
            // has no caller codec: retain the original image and report the filter.
            return false;
        }
    }

    private static OfficeDrawing CreateEmptySvgFilterResult(OfficeDrawing source, SvgInteractiveBounds bounds) {
        var result = new OfficeDrawing(source.Width, source.Height);
        result.AddClippedDrawingForRendering(new OfficeDrawing(source.Width, source.Height), 0D, 0D, OfficeClipPath.Empty(), 0D, 0D);
        ((OfficeDrawingGroup)result.Elements[0]).UnfilteredGeometryBounds = (bounds.Left, bounds.Top, bounds.Right, bounds.Bottom);
        return result;
    }

    private static float[] ResolveFilterPixels(int input, float[] graphic, float[] alpha, List<float[]> outputs) =>
        input == -1 ? graphic : input == -2 ? alpha : outputs[input];

    private static bool ContainsSvgFilterSemantics(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingText or OfficeDrawingRichText or OfficeDrawingLink) return true;
            if (element is OfficeDrawingImage image && image.AlternativeText != null) return true;
            if (element is OfficeDrawingGroup group &&
                (group.ActualText != null || ContainsSvgFilterSemantics(group.InnerDrawing))) return true;
            if (element is OfficeDrawingEffectGroup effect &&
                (ContainsSvgFilterSemantics(effect.InnerDrawing) ||
                (effect.SoftMask != null && ContainsSvgFilterSemantics(effect.SoftMask.InnerDrawing)))) return true;
            // Pattern cells may contain semantic content not present in the outer element list.
            if (element is OfficeDrawingTilingPattern) return true;
        }
        return false;
    }

    private static bool TryResolveFilterRegion(SvgFilterGraph graph, OfficeDrawing source,
        SvgInteractiveBounds box, OfficeTransform transform, out SvgInteractiveBounds region) {
        region = default;
        double bw = box.Right - box.Left, bh = box.Bottom - box.Top;
        if (!TryFilterRegionLength(graph.X, graph.UserRegion ? source.Width : bw, graph.UserRegion, out double x, out bool xp) ||
            !TryFilterRegionLength(graph.Y, graph.UserRegion ? source.Height : bh, graph.UserRegion, out double y, out bool yp) ||
            !TryFilterRegionLength(graph.Width, graph.UserRegion ? source.Width : bw, graph.UserRegion, out double w, out _) ||
            !TryFilterRegionLength(graph.Height, graph.UserRegion ? source.Height : bh, graph.UserRegion, out double h, out _)) return false;
        x += graph.UserRegion ? xp ? 0D : -graph.ViewX : box.Left;
        y += graph.UserRegion ? yp ? 0D : -graph.ViewY : box.Top;
        if (w < 0D || h < 0D || !FiniteFilterNumber(x + w) || !FiniteFilterNumber(y + h)) return false;
        double x1 = x * transform.M11 + transform.OffsetX, x2 = (x + w) * transform.M11 + transform.OffsetX;
        double y1 = y * transform.M22 + transform.OffsetY, y2 = (y + h) * transform.M22 + transform.OffsetY;
        if (!FiniteFilterNumber(x1) || !FiniteFilterNumber(x2) || !FiniteFilterNumber(y1) || !FiniteFilterNumber(y2)) return false;
        region = new SvgInteractiveBounds(Math.Min(x1, x2), Math.Min(y1, y2), Math.Max(x1, x2), Math.Max(y1, y2));
        return FiniteFilterNumber(region.Left) && FiniteFilterNumber(region.Top) &&
            FiniteFilterNumber(region.Right) && FiniteFilterNumber(region.Bottom) &&
            Math.Abs(region.Left) <= OfficeSvgDrawingReaderOptions.MaximumAllowedViewportDimension &&
            Math.Abs(region.Top) <= OfficeSvgDrawingReaderOptions.MaximumAllowedViewportDimension &&
            Math.Abs(region.Right) <= OfficeSvgDrawingReaderOptions.MaximumAllowedViewportDimension &&
            Math.Abs(region.Bottom) <= OfficeSvgDrawingReaderOptions.MaximumAllowedViewportDimension;
    }

    private static bool TryFilterRegionLength(string text, double basis, bool userSpace,
        out double length, out bool percentage) {
        if (userSpace || text.Trim().EndsWith("%", StringComparison.Ordinal))
            return TryViewportLength(text, basis, out length, out percentage);
        percentage = false;
        if (!TryParseFiniteNumber(text, 0D, out double fraction)) { length = 0D; return false; }
        length = fraction * basis;
        return FiniteFilterNumber(length);
    }

    private static void ClipFilterPixels(float[] pixels, int width, int height, SvgInteractiveBounds region,
        double left, double top, System.Threading.CancellationToken token) {
        for (int y = 0; y < height; y++) {
            token.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                if (left + x + 1D <= region.Left || left + x >= region.Right ||
                    top + y + 1D <= region.Top || top + y >= region.Bottom)
                    Array.Clear(pixels, (y * width + x) * 4, 4);
            }
        }
    }

    private static float[] ReadFilterPixels(OfficeRasterImage image, bool linear, System.Threading.CancellationToken token) {
        byte[] bytes = image.PixelBuffer;
        var pixels = new float[bytes.Length];
        for (int y = 0; y < image.Height; y++) {
            token.ThrowIfCancellationRequested();
            for (int i = y * image.Width * 4; i < (y + 1) * image.Width * 4; i += 4) {
                float alpha = bytes[i + 3] / 255F;
                for (int c = 0; c < 3; c++) {
                    double channel = bytes[i + c] / 255D;
                    pixels[i + c] = (float)(linear ? FilterSrgbToLinear(channel) : channel) * alpha;
                }
                pixels[i + 3] = alpha;
            }
        }
        return pixels;
    }

    private static OfficeRasterImage WriteFilterPixels(float[] pixels, int width, int height, bool linear,
        System.Threading.CancellationToken token) {
        var bytes = new byte[pixels.Length];
        for (int y = 0; y < height; y++) {
            token.ThrowIfCancellationRequested();
            for (int i = y * width * 4; i < (y + 1) * width * 4; i += 4) {
                double alpha = FilterClamp(pixels[i + 3]);
                bytes[i + 3] = FilterByte(alpha);
                if (alpha <= 0D) continue;
                for (int c = 0; c < 3; c++) {
                    double channel = FilterClamp(pixels[i + c] / alpha);
                    bytes[i + c] = FilterByte(linear ? FilterLinearToSrgb(channel) : channel);
                }
            }
        }
        return OfficeRasterImage.FromOwnedRgba32(width, height, bytes);
    }

    private static double FilterClamp(double value) => Math.Min(1D, Math.Max(0D, value));
    private static byte FilterByte(double value) => (byte)Math.Round(FilterClamp(value) * 255D);
    private static double FilterSrgbToLinear(double value) => value <= 0.04045D ? value / 12.92D : Math.Pow((value + 0.055D) / 1.055D, 2.4D);
    private static double FilterLinearToSrgb(double value) => value <= 0.0031308D ? value * 12.92D : 1.055D * Math.Pow(value, 1D / 2.4D) - 0.055D;
}

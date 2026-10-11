using OfficeIMO.Drawing;

namespace OfficeIMO.DjVu;

public sealed partial class DjVuPage {
    /// <summary>Decodes and composes this page in process, with bounded output and codec buffers.</summary>
    /// <remarks>Coordinates select an unrotated native region; resolution scaling precedes display rotation.
    /// Annotations are reported but are not painted. Unsupported image codecs fail explicitly.</remarks>
    public DjVuRenderResult Render(DjVuRenderOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var settings = (options ?? new DjVuRenderOptions()).Snapshot();
        var limits = Document.ReadOptions.Snapshot();
        limits.MaxCodecBytes = Math.Min(limits.MaxCodecBytes, settings.MaxBytes);
        var budget = new DjVuReadBudget(limits, cancellationToken);
        var region = settings.Region ?? new DjVuRectangle(0, 0, Width, Height);
        if (region.X < 0 || region.Y < 0 || region.Width <= 0 || region.Height <= 0 ||
            (long)region.X + region.Width > Width || (long)region.Y + region.Height > Height)
            throw new ArgumentOutOfRangeException(nameof(options), "The selected region must be inside the unrotated page.");
        int dpi = settings.Dpi ?? Dpi;
        long outputWidth = Math.Max(1, (long)Math.Round(region.Width * (double)dpi / Dpi, MidpointRounding.AwayFromZero));
        long outputHeight = Math.Max(1, (long)Math.Round(region.Height * (double)dpi / Dpi, MidpointRounding.AwayFromZero));
        if (outputWidth > int.MaxValue || outputHeight > int.MaxValue || outputWidth * outputHeight > settings.MaxPixels)
            throw new DjVuResourceLimitException(nameof(DjVuRenderOptions.MaxPixels));
        int width = (int)outputWidth, height = (int)outputHeight;
        budget.WorkingBytes((long)region.Width * region.Height * 4);
        var chunks = Document.PageChunks(Component, cancellationToken).ToList();
        var diagnostics = new List<OfficeConversionFidelityDiagnostic>();
        if (chunks.Any(c => c.Id == "ANTa" || c.Id == "ANTz"))
            diagnostics.Add(new OfficeConversionFidelityDiagnostic("djvu.render.annotation-omitted",
                "DjVu annotations and viewer settings are not painted.", OfficeConversionLossKind.Omission, "OfficeIMO.DjVu", "page " + Number));
        if (chunks.Any(c => (c.Id == "BG44" || c.Id == "FG44" || c.Id == "PM44" || c.Id == "BM44") &&
            c.Length >= 9 && c.Source[c.Offset] == 0 &&
            ((c.Source[c.Offset + 4] << 8 | c.Source[c.Offset + 5]) < 32 || (c.Source[c.Offset + 6] << 8 | c.Source[c.Offset + 7]) < 32)))
            diagnostics.Add(new OfficeConversionFidelityDiagnostic("djvu.render.short-iw44-edge",
                "Short IW44 image edges can differ from native decoder fixed-point reconstruction. See the qualified raster profile in SUPPORT.md.",
                OfficeConversionLossKind.Approximation, "OfficeIMO.DjVu", "page " + Number));
        OfficeRasterImage image = Compose(region, chunks, settings, budget);
        if (image.Width != width || image.Height != height) {
            bool simple = settings.Resampling == OfficeRasterResamplingMode.NearestNeighbor || settings.Resampling == OfficeRasterResamplingMode.Bilinear;
            long working;
            bool measured = simple
                ? OfficeRasterResampler.TryGetSimpleWorkingSetBytes(image.Width, image.Height, width, height, 0, out working)
                : OfficeRasterResampler.TryGetHighQualityWorkingSetBytes(image.Width, image.Height, width, height, settings.Resampling, out working);
            if (!measured) throw new DjVuResourceLimitException(nameof(DjVuRenderOptions.MaxBytes));
            budget.WorkingBytes(working);
            image = OfficeRasterResampler.Resize(image, width, height, settings.Resampling, OfficeRasterResamplingColorSpace.EncodedSrgb, cancellationToken);
        }
        int rotation = settings.ApplyRotation ? Rotation : 0;
        if (rotation != 0) {
            budget.WorkingBytes((long)image.Width * image.Height * 8);
            image = OfficeRasterTransforms.Rotate(image, rotation, settings.Background, cancellationToken);
        }
        return new DjVuRenderResult(image, dpi, region, rotation, diagnostics);
    }

    private OfficeRasterImage Compose(DjVuRectangle region, List<DjVuChunk> chunks, DjVuRenderOptions settings, DjVuReadBudget budget) {
        if (chunks.Any(c => c.Id == "BG2k" || c.Id == "FG2k"))
            throw new NotSupportedException("DjVu JPEG 2000 layers are not supported.");
        var backgroundChunks = chunks.Where(c => c.Id == "BG44" || c.Id == "BGjp" || c.Id == "PM44" || c.Id == "BM44").ToList();
        var foregroundChunks = chunks.Where(c => c.Id == "FG44" || c.Id == "FGjp").ToList();
        var palettes = chunks.Where(c => c.Id == "FGbz").ToList();
        bool hasMask = chunks.Any(c => c.Id == "Sjbz" || c.Id == "Smmr");
        if (!hasMask && backgroundChunks.Count == 0) throw new InvalidDataException("DjVu page has no supported image layer.");
        if (foregroundChunks.Count > 1 || palettes.Count > 1 || foregroundChunks.Count != 0 && palettes.Count != 0 ||
            !hasMask && (foregroundChunks.Count != 0 || palettes.Count != 0) ||
            palettes.Count != 0 && chunks.Any(c => c.Id == "Smmr"))
            throw new InvalidDataException("Conflicting DjVu foreground layers.");
        if (backgroundChunks.Any(c => c.Id == "BGjp") && backgroundChunks.Count != 1) throw new InvalidDataException("Conflicting DjVu background codecs.");
        Jb2Image? mask = hasMask ? DecodeMask(budget) : null;
        if (mask != null) budget.RetainBytes(mask.RetainedBytes);
        var background = backgroundChunks.Count == 0 ? null : DjVuColorLayer.Decode(backgroundChunks, this, budget);
        var foreground = foregroundChunks.Count == 0 ? null : DjVuColorLayer.Decode(foregroundChunks, this, budget);
        var palette = palettes.Count == 0 ? null : DjVuPalette.Decode(palettes[0], mask!.Placements.Count, budget);
        budget.WorkingBytes((long)region.Width * region.Height * 4);
        var output = new byte[checked(region.Width * region.Height * 4)];
        for (int row = 0; row < region.Height; row++) {
            budget.Cancellation.ThrowIfCancellationRequested();
            int y = region.Y + region.Height - 1 - row;
            for (int col = 0; col < region.Width; col++) {
                int offset = (row * region.Width + col) * 4;
                if (background != null) background.Paint(region.X + col, y, output, offset);
                else { output[offset] = settings.Background.R; output[offset + 1] = settings.Background.G; output[offset + 2] = settings.Background.B; output[offset + 3] = 255; }
            }
        }
        if (mask != null) PaintMask(mask, foreground, palette, region, output, settings.MaxMaskPaintSamples, budget.Cancellation);
        if (settings.Gamma != Gamma) {
            var correction = new byte[256];
            for (int i = 0; i < correction.Length; i++) correction[i] = (byte)Math.Round(255 * Math.Pow(i / 255.0, Gamma / settings.Gamma));
            for (int i = 0; i < output.Length; i += 4) {
                if ((i & 16383) == 0) budget.Cancellation.ThrowIfCancellationRequested();
                output[i] = correction[output[i]]; output[i + 1] = correction[output[i + 1]]; output[i + 2] = correction[output[i + 2]];
            }
        }
        // Codec objects are local to composition. Subsequent resampling accounts for the returned raster only.
        budget.ReleaseBytes(budget.RetainedBytes);
        return OfficeRasterImage.FromOwnedRgba32(region.Width, region.Height, output);
    }

    private static void PaintMask(Jb2Image mask, DjVuColorLayer? foreground, DjVuPalette? palette, DjVuRectangle region, byte[] output, long maxSamples, CancellationToken cancellation) {
        long samples = 0;
        for (int i = 0; i < mask.Placements.Count; i++) {
            cancellation.ThrowIfCancellationRequested();
            var placement = mask.Placements[i];
            int left = (int)Math.Max(0L, Math.Min(placement.Bitmap.Width, (long)region.X - placement.X));
            int right = (int)Math.Max(0L, Math.Min(placement.Bitmap.Width, (long)region.X + region.Width - placement.X));
            int bottom = (int)Math.Max(0L, Math.Min(placement.Bitmap.Height, (long)region.Y - placement.Y));
            int top = (int)Math.Max(0L, Math.Min(placement.Bitmap.Height, (long)region.Y + region.Height - placement.Y));
            if (left == right || bottom == top) continue;
            long count = (long)(right - left) * (top - bottom);
            if (count > maxSamples - samples) throw new DjVuResourceLimitException(nameof(DjVuRenderOptions.MaxMaskPaintSamples));
            samples += count;
            for (int y = bottom; y < top; y++) {
                cancellation.ThrowIfCancellationRequested();
                for (int x = left; x < right; x++) {
                    if (placement.Bitmap.Pixels[y * placement.Bitmap.Width + x] == 0) continue;
                    int pageX = placement.X + x, pageY = placement.Y + y;
                    int offset = ((region.Y + region.Height - 1 - pageY) * region.Width + pageX - region.X) * 4;
                    if (palette != null) palette.Paint(i, output, offset);
                    else if (foreground != null) foreground.PaintForeground(pageX, pageY, output, offset);
                    else { output[offset] = output[offset + 1] = output[offset + 2] = 0; output[offset + 3] = 255; }
                }
            }
        }
    }
}

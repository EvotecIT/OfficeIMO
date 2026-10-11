using OfficeIMO.Drawing;

namespace OfficeIMO.DjVu;

internal sealed class DjVuColorLayer {
    private readonly OfficeRasterImage _image;
    private readonly int _reduction;
    internal DjVuColorLayer(OfficeRasterImage image, int reduction) { _image = image; _reduction = reduction; }

    internal void Paint(int x, int y, byte[] output, int offset) {
        if (_reduction == 1) {
            Buffer.BlockCopy(_image.PixelBuffer, ((_image.Height - 1 - y) * _image.Width + x) * 4, output, offset, 4);
            return;
        }
        // The DjVu background grid uses sixteenth-pixel coordinates, rounded up.
        // Interpolate vertically first, rounding each pass to an eight-bit sample.
        // This is a format rule, rather than a generic image resize operation.
        int sx = (x * 16 + 8 + _reduction - 1) / _reduction - 8;
        int sy = (y * 16 + 8 + _reduction - 1) / _reduction - 8;
        int x0 = sx >> 4, y0 = sy >> 4, fx = sx & 15, fy = sy & 15;
        int xa = Math.Max(0, Math.Min(_image.Width - 1, x0));
        int xb = Math.Max(0, Math.Min(_image.Width - 1, x0 + 1));
        int ya = _image.Height - 1 - Math.Max(0, Math.Min(_image.Height - 1, y0));
        int yb = _image.Height - 1 - Math.Max(0, Math.Min(_image.Height - 1, y0 + 1));
        var pixels = _image.PixelBuffer;
        for (int c = 0; c < 3; c++) {
            int left = (pixels[(ya * _image.Width + xa) * 4 + c] * (16 - fy) + pixels[(yb * _image.Width + xa) * 4 + c] * fy + 8) >> 4;
            int right = (pixels[(ya * _image.Width + xb) * 4 + c] * (16 - fy) + pixels[(yb * _image.Width + xb) * 4 + c] * fy + 8) >> 4;
            output[offset + c] = (byte)((left * (16 - fx) + right * fx + 8) >> 4);
        }
        output[offset + 3] = 255;
    }

    internal void PaintForeground(int x, int y, byte[] output, int offset) {
        // Foreground colour is a per-mask-sample stencil, without background interpolation.
        int row = _image.Height - 1 - y / _reduction, column = x / _reduction;
        Buffer.BlockCopy(_image.PixelBuffer, (row * _image.Width + column) * 4, output, offset, 4);
    }

    internal static DjVuColorLayer Decode(IReadOnlyList<DjVuChunk> chunks, DjVuPage page, DjVuReadBudget budget) {
        OfficeRasterImage image;
        if (chunks[0].Id == "BGjp" || chunks[0].Id == "FGjp") {
            if (chunks.Count != 1) throw new InvalidDataException("Multiple JPEG chunks in a DjVu layer.");
            var chunk = chunks[0];
            if (chunk.Length < 2 || chunk.Source[chunk.Offset] != 255 || chunk.Source[chunk.Offset + 1] != 216)
                throw new InvalidDataException("DjVu JPEG layer does not contain a JPEG codestream.");
            budget.WorkingBytes(chunk.Length + 64L * 1024);
            byte[] data = new byte[chunk.Length];
            Buffer.BlockCopy(chunk.Source, chunk.Offset, data, 0, data.Length);
            if (!OfficeImageReader.TryIdentifyByContent(data, null, budget.Cancellation, out var header) ||
                header.Format != OfficeImageFormat.Jpeg || header.Width <= 0 || header.Height <= 0)
                throw new InvalidDataException("Invalid DjVu JPEG layer header.");
            long pixels = (long)header.Width * header.Height;
            if (pixels > budget.Options.MaxPagePixels)
                throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxPagePixels));
            budget.WorkingBytes(data.LongLength + 64L * 1024 + pixels * 4);
            var decodeOptions = new OfficeRasterDecodeOptions {
                CancellationToken = budget.Cancellation, MaximumDecodedPixels = Math.Min(budget.Options.MaxPagePixels, 50_000_000),
                MaximumEncodedBytes = Math.Min(128 * 1024 * 1024, Math.Max(1, chunk.Length)),
                // Reserve the part of Core's fixed codec ceiling unavailable to
                // this operation, so its allocation guards enforce our lower cap.
                RetainedManagedBytes = budget.RetainedBytes + data.Length + Math.Max(0, OfficeRasterGuards.MaximumDecodedBytes - budget.Options.MaxCodecBytes),
                ApplyExifOrientation = false
            };
            if (!OfficeRasterImageDecoder.TryDecode(data, decodeOptions, out var decoded, out var info) || decoded == null)
                throw new InvalidDataException("Invalid DjVu JPEG layer: " + info.Diagnostic);
            image = decoded;
            budget.RetainBytes(image.PixelBuffer.LongLength);
        } else {
            var decoder = new Iw44Decoder(budget);
            foreach (var chunk in chunks) decoder.Read(chunk);
            byte[] rgb = decoder.Reconstruct();
            budget.WorkingBytes(decoder.RetainedBytes + rgb.LongLength + (long)decoder.Width * decoder.Height * 4);
            byte[] rgba = new byte[checked(decoder.Width * decoder.Height * 4)];
            for (int i = 0, j = 0; i < rgba.Length; i += 4, j += 3) {
                if ((i & 16383) == 0) budget.Cancellation.ThrowIfCancellationRequested();
                rgba[i] = rgb[j]; rgba[i + 1] = rgb[j + 1]; rgba[i + 2] = rgb[j + 2]; rgba[i + 3] = 255;
            }
            image = OfficeRasterImage.FromOwnedRgba32(decoder.Width, decoder.Height, rgba);
            budget.RetainBytes(rgba.LongLength);
        }
        int reduction = 0;
        for (int d = 1; d <= 12; d++) {
            if ((page.Width + d - 1) / d == image.Width && (page.Height + d - 1) / d == image.Height) { reduction = d; break; }
        }
        if (reduction == 0) throw new InvalidDataException("DjVu layer dimensions do not match the page's subsampling grid.");
        return new DjVuColorLayer(image, reduction);
    }
}

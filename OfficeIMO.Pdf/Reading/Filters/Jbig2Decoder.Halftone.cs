using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf.Filters;

internal static partial class Jbig2Decoder {
    private sealed partial class DecodingContext {
        private readonly Dictionary<int, List<Bitmap>> _patterns = new();

        // Pattern dictionary segment (7.4.4) and decoding procedure (6.7.5).
        private void ProcessPatternDictionary(Segment segment) {
            int position = segment.DataStart;
            int flags = ReadByte(segment.Data, ref position, segment.DataEnd);
            bool mmr = (flags & 1) != 0;
            int template = (flags >> 1) & 3;
            int patternWidth = ReadByte(segment.Data, ref position, segment.DataEnd);
            int patternHeight = ReadByte(segment.Data, ref position, segment.DataEnd);
            uint grayMaximum = ReadUInt32(segment.Data, ref position, segment.DataEnd);
            if (patternWidth == 0 || patternHeight == 0 || grayMaximum >= MaximumSymbols) throw new InvalidDataException("JBIG2 pattern dictionary is invalid.");
            int count = (int)grayMaximum + 1;
            long collectiveWidth = (long)count * patternWidth;
            if (collectiveWidth * patternHeight > MaximumBitmapPixels) throw new InvalidDataException("JBIG2 pattern dictionary is too large.");
            Bitmap collective;
            if (mmr) {
                collective = DecodeMmrAt(segment.Data, position, segment.DataEnd, (int)collectiveWidth, patternHeight, out _);
            } else {
                var decoder = new ArithmeticDecoder(segment.Data, position, segment.DataEnd, Charge);
                int[] at = { -patternWidth, 0, -3, -1, 2, -2, -2, -2 };
                collective = DecodeGeneric(decoder, CreateContexts(1 << 16), (int)collectiveWidth, patternHeight, template, false, at);
            }
            List<Bitmap> patterns = CreateSymbolList(count);
            for (int index = 0; index < count; index++) {
                Bitmap pattern = CreateBitmap(patternWidth, patternHeight);
                pattern.Compose(collective, -index * patternWidth, 0, 4);
                patterns.Add(pattern);
            }
            _patterns[segment.Number] = patterns;
        }

        // Halftone region segment (7.4.5) and decoding procedure (6.6.5).
        private void ProcessHalftoneRegion(Segment segment) {
            int position = segment.DataStart;
            var info = ReadRegionInfo(segment, ref position);
            int flags = ReadByte(segment.Data, ref position, segment.DataEnd);
            bool mmr = (flags & 1) != 0;
            int template = (flags >> 1) & 3;
            bool enableSkip = (flags & 8) != 0;
            int combination = (flags >> 4) & 7;
            int defaultPixel = (flags >> 7) & 1;
            uint gridWidthValue = ReadUInt32(segment.Data, ref position, segment.DataEnd);
            uint gridHeightValue = ReadUInt32(segment.Data, ref position, segment.DataEnd);
            int gridX = (int)ReadUInt32(segment.Data, ref position, segment.DataEnd);
            int gridY = (int)ReadUInt32(segment.Data, ref position, segment.DataEnd);
            int vectorX = (ReadByte(segment.Data, ref position, segment.DataEnd) << 8) | ReadByte(segment.Data, ref position, segment.DataEnd);
            int vectorY = (ReadByte(segment.Data, ref position, segment.DataEnd) << 8) | ReadByte(segment.Data, ref position, segment.DataEnd);
            if (gridWidthValue > int.MaxValue || gridHeightValue > int.MaxValue ||
                (long)gridWidthValue * gridHeightValue > MaximumBitmapPixels) throw new InvalidDataException("JBIG2 halftone grid is too large.");
            int gridWidth = (int)gridWidthValue, gridHeight = (int)gridHeightValue;
            List<Bitmap>? patterns = null;
            foreach (int referred in segment.Referred) {
                if (_patterns.TryGetValue(referred, out List<Bitmap>? found)) {
                    patterns = found;
                    break;
                }
            }
            if (patterns == null || patterns.Count == 0) throw new InvalidDataException("JBIG2 halftone region lacks a pattern dictionary.");
            Bitmap region = CreateBitmap(info.Width, info.Height);
            if (defaultPixel != 0) region.Fill(1);
            int patternWidth = patterns[0].Width, patternHeight = patterns[0].Height;
            Bitmap? skip = null;
            if (enableSkip) {
                skip = CreateBitmap(gridWidth, gridHeight);
                for (int m = 0; m < gridHeight; m++) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    for (int n = 0; n < gridWidth; n++) {
                        long x = ((long)gridX + (long)m * vectorY + (long)n * vectorX) >> 8;
                        long y = ((long)gridY + (long)m * vectorX - (long)n * vectorY) >> 8;
                        if (x + patternWidth <= 0 || x >= info.Width || y + patternHeight <= 0 || y >= info.Height) skip.Set(n, m, 1);
                    }
                }
            }
            int bitsPerValue = CodeLength(patterns.Count);
            int[] gray = DecodeGrayScaleImage(segment, position, mmr, template, gridWidth, gridHeight, bitsPerValue, skip);
            Charge((long)gridWidth * gridHeight);
            for (int m = 0; m < gridHeight; m++) {
                _cancellationToken.ThrowIfCancellationRequested();
                for (int n = 0; n < gridWidth; n++) {
                    long x = ((long)gridX + (long)m * vectorY + (long)n * vectorX) >> 8;
                    long y = ((long)gridY + (long)m * vectorX - (long)n * vectorY) >> 8;
                    if (x + patternWidth <= 0 || x >= info.Width || y + patternHeight <= 0 || y >= info.Height) continue;
                    int value = Math.Min(gray[m * gridWidth + n], patterns.Count - 1);
                    region.Compose(patterns[value], (int)x, (int)y, combination);
                }
            }
            Place(segment, region, info);
        }

        // Gray-scale image decoding procedure (Annex C.5): Gray-coded bit-planes, most significant first.
        private int[] DecodeGrayScaleImage(Segment segment, int position, bool mmr, int template, int width, int height, int bits, Bitmap? skip) {
            ReserveBytes((long)width * height * sizeof(int));
            var values = new int[(long)width * height];
            if (bits == 0 || width == 0 || height == 0) return values;
            var planes = new Bitmap[bits];
            ArithmeticDecoder? decoder = null;
            byte[]? contexts = null;
            if (!mmr) {
                decoder = new ArithmeticDecoder(segment.Data, position, segment.DataEnd, Charge);
                contexts = CreateContexts(1 << 16);
            }
            for (int plane = bits - 1; plane >= 0; plane--) {
                if (mmr) {
                    planes[plane] = DecodeMmrAt(segment.Data, position, segment.DataEnd, width, height, out int consumed);
                    position += consumed;
                } else {
                    planes[plane] = DecodeGeneric(decoder!, contexts!, width, height, template, false, Array.Empty<int>(), skip);
                }
            }
            for (int plane = bits - 2; plane >= 0; plane--) planes[plane].Compose(planes[plane + 1], 0, 0, 2);
            Charge((long)width * height * bits);
            for (int y = 0; y < height; y++) {
                _cancellationToken.ThrowIfCancellationRequested();
                for (int x = 0; x < width; x++) {
                    int value = 0;
                    for (int plane = bits - 1; plane >= 0; plane--) value = (value << 1) | planes[plane].Get(x, y);
                    values[y * width + x] = value;
                }
            }
            return values;
        }

        private Bitmap DecodeMmrAt(byte[] data, int position, int end, int width, int height, out int consumed) {
            consumed = 0;
            var region = CreateBitmap(width, height);
            if (width == 0 || height == 0) return region;
            int length = Math.Max(0, end - position);
            ReserveBytes(length + ((long)width + 7) / 8 * height);
            var encoded = new byte[length];
            Buffer.BlockCopy(data, position, encoded, 0, length);
            byte[] rows = OfficeFaxDecoder.DecodeMmr(encoded, width, height, checked((int)(((long)width + 7) / 8 * height)), _cancellationToken, out consumed);
            CopyPackedRows(rows, region, width, height);
            return region;
        }

        private void CopyPackedRows(byte[] rows, Bitmap region, int width, int height) {
            int stride = (width + 7) / 8;
            for (int y = 0; y < height; y++) {
                _cancellationToken.ThrowIfCancellationRequested();
                for (int x = 0; x < width; x++) region.Set(x, y, (rows[y * stride + x / 8] >> (7 - (x & 7))) & 1);
            }
        }
    }
}

using OfficeIMO.Drawing;

namespace OfficeIMO.DjVu;

// DjVu v3 section 8.3.15. The owned Core T.6 codec also serves TIFF and PDF.
internal static class DjVuMmrDecoder {
    internal static Jb2Image Decode(DjVuChunk chunk, DjVuReadBudget budget) {
        byte[] source = chunk.Source;
        int position = chunk.Offset, end = position + chunk.Length;
        if (end - position < 8 || source[position++] != 'M' || source[position++] != 'M' || source[position++] != 'R')
            throw new InvalidDataException("Invalid DjVu MMR header.");
        int flags = source[position++];
        if ((flags & ~3) != 0) throw new NotSupportedException("Unsupported DjVu MMR flags.");
        int width = DjVuBinary.U16(source, position), height = DjVuBinary.U16(source, position + 2); position += 4;
        if (width == 0 || height == 0) throw new InvalidDataException("Empty DjVu MMR image.");
        if ((long)width * height > budget.Options.MaxPagePixels) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxPagePixels));
        int rowsPerStripe = height;
        if ((flags & 2) != 0) {
            if (end - position < 2) throw new InvalidDataException("Truncated DjVu MMR stripe header.");
            rowsPerStripe = DjVuBinary.U16(source, position); position += 2;
            if (rowsPerStripe == 0) throw new InvalidDataException("Empty DjVu MMR stripe.");
        }
        int stride = (width + 7) / 8;
        budget.WorkingBytes((long)width * height + (long)stride * Math.Min(height, rowsPerStripe) + chunk.Length);
        var pixels = new byte[checked(width * height)];
        for (int row = 0; row < height; row += rowsPerStripe) {
            budget.Cancellation.ThrowIfCancellationRequested();
            int rows = Math.Min(rowsPerStripe, height - row), size = end - position;
            if ((flags & 2) != 0) {
                if (end - position < 4) throw new InvalidDataException("Truncated DjVu MMR stripe length.");
                uint length = DjVuBinary.U32(source, position); position += 4;
                if (length > end - position) throw new InvalidDataException("DjVu MMR stripe exceeds its chunk.");
                size = (int)length;
            }
            byte[] encoded = new byte[size];
            Buffer.BlockCopy(source, position, encoded, 0, size); position += size;
            byte[] packed = OfficeFaxDecoder.Decode(encoded, width, rows, -1, endOfLine: false, byteAligned: false,
                blackIsOne: (flags & 1) == 0, endOfBlock: false, maximumBytes: checked(stride * rows), cancellationToken: budget.Cancellation);
            for (int y = 0; y < rows; y++) {
                budget.Cancellation.ThrowIfCancellationRequested();
                for (int x = 0; x < width; x++) pixels[(height - 1 - row - y) * width + x] = (byte)((packed[y * stride + x / 8] >> (7 - x % 8)) & 1);
            }
        }
        if (position != end) throw new InvalidDataException("Trailing DjVu MMR stripe bytes.");
        var bitmap = new Jb2Bitmap(width, height, pixels);
        return new Jb2Image(width, height, new List<Jb2Bitmap>(), new List<Jb2Placement> { new Jb2Placement(bitmap, 0, 0) });
    }
}

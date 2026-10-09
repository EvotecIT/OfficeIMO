namespace OfficeIMO.DjVu;

internal sealed class Jb2Bitmap {
    internal readonly int Width, Height;
    // Unrotated, bottom-up samples: 0 = white, 1 = black.
    internal readonly byte[] Pixels;
    internal Jb2Bitmap(int width, int height, byte[] pixels) { Width = width; Height = height; Pixels = pixels; }
    internal int At(int x, int y) => (uint)x < (uint)Width && (uint)y < (uint)Height ? Pixels[y * Width + x] : 0;

    internal Jb2Bitmap Trim(CancellationToken cancellation) {
        cancellation.ThrowIfCancellationRequested();
        if (Pixels.Length == 0) return new Jb2Bitmap(0, 0, Array.Empty<byte>());
        int left = Width, right = -1, bottom = Height, top = -1;
        for (int y = 0; y < Height; y++) {
            cancellation.ThrowIfCancellationRequested();
            for (int x = 0; x < Width; x++) {
                if (Pixels[y * Width + x] == 0) continue;
                left = Math.Min(left, x); right = Math.Max(right, x);
                bottom = Math.Min(bottom, y); top = Math.Max(top, y);
            }
        }
        if (right < 0) return new Jb2Bitmap(0, 0, Array.Empty<byte>());
        if (left == 0 && bottom == 0 && right == Width - 1 && top == Height - 1) return this;
        int width = right - left + 1, height = top - bottom + 1;
        var pixels = new byte[width * height];
        for (int y = 0; y < height; y++) Array.Copy(Pixels, (y + bottom) * Width + left, pixels, y * width, width);
        return new Jb2Bitmap(width, height, pixels);
    }
}

internal sealed class Jb2Placement {
    internal readonly Jb2Bitmap Bitmap;
    internal readonly int X, Y;
    internal Jb2Placement(Jb2Bitmap bitmap, int x, int y) { Bitmap = bitmap; X = x; Y = y; }
}

internal sealed class Jb2Image {
    internal readonly int Width, Height;
    internal readonly IReadOnlyList<Jb2Bitmap> Library;
    internal readonly IReadOnlyList<Jb2Placement> Placements;
    internal long RetainedBytes {
        get {
            var bitmaps = new HashSet<Jb2Bitmap>(Library);
            foreach (var placement in Placements) bitmaps.Add(placement.Bitmap);
            return bitmaps.Sum(b => b.Pixels.LongLength + 32L) + Library.Count * 8L + Placements.Count * 40L;
        }
    }
    internal Jb2Image(int width, int height, List<Jb2Bitmap> library, List<Jb2Placement> placements) {
        Width = width; Height = height; Library = library.AsReadOnly(); Placements = placements.AsReadOnly();
    }
}

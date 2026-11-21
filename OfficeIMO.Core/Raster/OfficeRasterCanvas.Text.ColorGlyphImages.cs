namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    internal const int MaximumCachedColorGlyphImages = 128;
    internal const long MaximumCachedColorGlyphPixelBytes = 16L * 1024 * 1024;

    // Decoded PNG glyphs, least recently used first. The key is the font's shared PNG array, so a glyph the
    // font cache evicts and reads again decodes again, and this cache drops the old entry in its turn.
    private LinkedList<(byte[] Png, OfficeRasterImage? Image, long Bytes)>? _colorGlyphImageOrder;
    private Dictionary<byte[], LinkedListNode<(byte[] Png, OfficeRasterImage? Image, long Bytes)>>? _colorGlyphImages;
    private long _colorGlyphImageBytes;

    /// <summary>The number of decoded color glyph images this canvas retains.</summary>
    internal int CachedColorGlyphImageCount => _colorGlyphImages?.Count ?? 0;

    /// <summary>The decoded pixel bytes of the color glyph images this canvas retains.</summary>
    internal long CachedColorGlyphImageBytes => _colorGlyphImageBytes;

    // A COLR layer is filled; a CBDT or sbix PNG layer is drawn through its box, which arrives as one
    // contour (top-left, top-right, bottom-right, bottom-left) so text scaling, slant and rotation move it.
    private void PaintColorGlyphLayer(OfficeColorGlyphContours layer, double boldOffset) {
        if (layer.Bitmap == null) {
            FillTextContours(layer.Contours, layer.Color, boldOffset);
            return;
        }
        IReadOnlyList<List<OfficePoint>> painted = TransformPositionedGlyphContours(layer.Contours);
        if (_textInkObserver != null) {
            _textInkObserver(painted);
            return;
        }
        if (painted.Count == 0 || painted[0].Count < 4) return;
        OfficeRasterImage? image = DecodeColorGlyphImage(layer.Bitmap.Png);
        if (image == null || image.Width <= 0 || image.Height <= 0) return;
        OfficePoint origin = painted[0][0], right = painted[0][1], down = painted[0][3];
        var transform = new OfficeTransform(
            (right.X - origin.X) / image.Width, (right.Y - origin.Y) / image.Width,
            (down.X - origin.X) / image.Height, (down.Y - origin.Y) / image.Height,
            origin.X, origin.Y);
        DrawAffineImage(image, transform, 1D, interpolate: true);
    }

    // A glyph's PNG is shared by every use of the glyph, so it is decoded once and kept while recently used,
    // within both an entry and a decoded-pixel budget.
    private OfficeRasterImage? DecodeColorGlyphImage(byte[] png) {
        _colorGlyphImages ??= new Dictionary<byte[], LinkedListNode<(byte[] Png, OfficeRasterImage? Image, long Bytes)>>(SameArray.Instance);
        _colorGlyphImageOrder ??= new LinkedList<(byte[] Png, OfficeRasterImage? Image, long Bytes)>();
        if (_colorGlyphImages.TryGetValue(png, out var cached)) {
            _colorGlyphImageOrder.Remove(cached);
            _colorGlyphImageOrder.AddLast(cached);
            return cached.Value.Image;
        }
        OfficeRasterImage? image;
        try {
            image = OfficeRasterImageDecoder.Decode(png);
        } catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or ArgumentException) {
            image = null;
        }
        long bytes = image == null ? 0L : (long)image.Width * image.Height * 4L;
        // An image larger than the whole budget is drawn but not kept.
        if (bytes > MaximumCachedColorGlyphPixelBytes) return image;
        while (_colorGlyphImageOrder.First != null &&
            (_colorGlyphImages.Count >= MaximumCachedColorGlyphImages || _colorGlyphImageBytes + bytes > MaximumCachedColorGlyphPixelBytes)) {
            var oldest = _colorGlyphImageOrder.First;
            _colorGlyphImageOrder.RemoveFirst();
            _colorGlyphImages.Remove(oldest.Value.Png);
            _colorGlyphImageBytes -= oldest.Value.Bytes;
        }
        _colorGlyphImages[png] = _colorGlyphImageOrder.AddLast((png, image, bytes));
        _colorGlyphImageBytes += bytes;
        return image;
    }

    private sealed class SameArray : IEqualityComparer<byte[]> {
        internal static readonly SameArray Instance = new();
        public bool Equals(byte[]? x, byte[]? y) => ReferenceEquals(x, y);
        public int GetHashCode(byte[] obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
    }
}

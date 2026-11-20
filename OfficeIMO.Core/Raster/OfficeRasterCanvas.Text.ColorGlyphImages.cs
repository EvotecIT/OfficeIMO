namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private Dictionary<byte[], OfficeRasterImage?>? _colorGlyphImages;

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

    // A glyph's PNG is shared by every use of the glyph, so decode it once per canvas.
    private OfficeRasterImage? DecodeColorGlyphImage(byte[] png) {
        _colorGlyphImages ??= new Dictionary<byte[], OfficeRasterImage?>(SameArray.Instance);
        if (_colorGlyphImages.TryGetValue(png, out OfficeRasterImage? cached)) return cached;
        OfficeRasterImage? image;
        try {
            image = OfficeRasterImageDecoder.Decode(png);
        } catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or ArgumentException) {
            image = null;
        }
        _colorGlyphImages[png] = image;
        return image;
    }

    private sealed class SameArray : IEqualityComparer<byte[]> {
        internal static readonly SameArray Instance = new();
        public bool Equals(byte[]? x, byte[]? y) => ReferenceEquals(x, y);
        public int GetHashCode(byte[] obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
    }
}

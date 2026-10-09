namespace OfficeIMO.OpenDocument;

/// <summary>A positive integer-sized local coordinate canvas for native ODF paths and polygons.</summary>
public readonly struct OdfViewBox {
    /// <summary>Creates a local canvas. Coordinates are unitless; the shape's bounds determine its page size.</summary>
    public OdfViewBox(int x, int y, int width, int height) {
        if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        X = x; Y = y; Width = width; Height = height;
    }
    /// <summary>Local horizontal origin.</summary>
    public int X { get; }
    /// <summary>Local vertical origin.</summary>
    public int Y { get; }
    /// <summary>Local canvas width.</summary>
    public int Width { get; }
    /// <summary>Local canvas height.</summary>
    public int Height { get; }
    /// <inheritdoc />
    public override string ToString() => string.Join(" ", new[] { X, Y, Width, Height }.Select(value => value.ToString(CultureInfo.InvariantCulture)));
    internal void Validate() {
        if (Width <= 0 || Height <= 0) throw new ArgumentException("A view box must have positive width and height.");
    }
    internal static OdfViewBox Parse(string? value) {
        if (value == null || value.Length > 128) throw new InvalidDataException("The shape requires a supported svg:viewBox.");
        string[] numbers = value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (numbers.Length != 4) throw new InvalidDataException("A view box requires four integer coordinates.");
        int[] values = numbers.Select(number => int.Parse(number, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture)).ToArray();
        return new OdfViewBox(values[0], values[1], values[2], values[3]);
    }
}

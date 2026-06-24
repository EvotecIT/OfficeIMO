using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a source-image crop rectangle using fractions of the original image edges.
/// </summary>
public sealed class PdfImageSourceCrop {
    private double _left;
    private double _top;
    private double _right;
    private double _bottom;

    /// <summary>Creates a crop rectangle with no cropped edges.</summary>
    public PdfImageSourceCrop() {
    }

    /// <summary>Creates a crop rectangle using fractions from 0 to less than 1 for each edge.</summary>
    public PdfImageSourceCrop(double left, double top, double right, double bottom) {
        Set(left, top, right, bottom);
    }

    /// <summary>Fraction cropped from the left edge of the source image.</summary>
    public double Left {
        get => _left;
        set => Set(value, _top, _right, _bottom);
    }

    /// <summary>Fraction cropped from the top edge of the source image.</summary>
    public double Top {
        get => _top;
        set => Set(_left, value, _right, _bottom);
    }

    /// <summary>Fraction cropped from the right edge of the source image.</summary>
    public double Right {
        get => _right;
        set => Set(_left, _top, value, _bottom);
    }

    /// <summary>Fraction cropped from the bottom edge of the source image.</summary>
    public double Bottom {
        get => _bottom;
        set => Set(_left, _top, _right, value);
    }

    internal bool HasCrop => Left > 0D || Top > 0D || Right > 0D || Bottom > 0D;

    internal PdfImageSourceCrop Clone() => new(Left, Top, Right, Bottom);

    internal OfficeImageSourceCrop ToOfficeImageSourceCrop() => new(Left, Top, Right, Bottom);

    private void Set(double left, double top, double right, double bottom) {
        OfficeImageSourceCrop crop = OfficeImageSourceCrop.FromStrictFractions(left, top, right, bottom);
        _left = crop.Left;
        _top = crop.Top;
        _right = crop.Right;
        _bottom = crop.Bottom;
    }
}

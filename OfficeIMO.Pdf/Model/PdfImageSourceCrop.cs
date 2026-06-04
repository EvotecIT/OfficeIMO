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

    private void Set(double left, double top, double right, double bottom) {
        ValidateCropFraction(left, nameof(Left));
        ValidateCropFraction(top, nameof(Top));
        ValidateCropFraction(right, nameof(Right));
        ValidateCropFraction(bottom, nameof(Bottom));
        if (left + right >= 1D) {
            throw new System.ArgumentOutOfRangeException(nameof(left), "Image source crop left and right fractions must leave a visible source width.");
        }

        if (top + bottom >= 1D) {
            throw new System.ArgumentOutOfRangeException(nameof(top), "Image source crop top and bottom fractions must leave a visible source height.");
        }

        _left = left;
        _top = top;
        _right = right;
        _bottom = bottom;
    }

    private static void ValidateCropFraction(double value, string paramName) {
        if (value < 0D || value >= 1D || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentOutOfRangeException(paramName, "Image source crop fractions must be finite values from 0 to less than 1.");
        }
    }
}

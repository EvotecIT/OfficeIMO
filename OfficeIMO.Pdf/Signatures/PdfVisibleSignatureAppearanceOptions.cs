using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Visible signature widget and dependency-free appearance-stream settings.</summary>
public sealed class PdfVisibleSignatureAppearanceOptions {
    private int _pageNumber = 1;
    private double _width = 180;
    private double _height = 48;
    private double _fontSize = 10;
    private double _imagePadding = 4;
    private byte[]? _imageBytes;

    /// <summary>One-based page number that receives the signature widget.</summary>
    public int PageNumber {
        get => _pageNumber;
        set => _pageNumber = value > 0 ? value : throw new ArgumentOutOfRangeException(nameof(value), "Page number must be positive.");
    }

    /// <summary>Left edge in PDF points.</summary>
    public double X { get; set; } = 36;

    /// <summary>Bottom edge in PDF points.</summary>
    public double Y { get; set; } = 36;

    /// <summary>Widget width in PDF points.</summary>
    public double Width {
        get => _width;
        set => _width = ValidatePositive(value, nameof(value));
    }

    /// <summary>Widget height in PDF points.</summary>
    public double Height {
        get => _height;
        set => _height = ValidatePositive(value, nameof(value));
    }

    /// <summary>Appearance text. Defaults to the signature field name when omitted.</summary>
    public string? Text { get; set; }

    /// <summary>Whether the appearance text is drawn. Defaults to true for compatibility.</summary>
    public bool ShowText { get; set; } = true;

    /// <summary>Helvetica appearance font size in points.</summary>
    public double FontSize {
        get => _fontSize;
        set => _fontSize = ValidatePositive(value, nameof(value));
    }

    /// <summary>Appearance background color.</summary>
    public PdfColor BackgroundColor { get; set; } = PdfColor.White;

    /// <summary>Appearance border color.</summary>
    public PdfColor BorderColor { get; set; } = PdfColor.Gray;

    /// <summary>Appearance text color.</summary>
    public PdfColor TextColor { get; set; } = PdfColor.Black;

    /// <summary>
    /// Optional raster image drawn behind the appearance text. The supplied bytes are defensively copied
    /// and accept the same managed raster formats as <see cref="PdfDocument.TryValidateImageBytes"/>.
    /// </summary>
    public byte[]? ImageBytes {
        get => _imageBytes is null ? null : (byte[])_imageBytes.Clone();
        set => _imageBytes = value is null ? null : (byte[])value.Clone();
    }

    /// <summary>How the image is fitted into the padded signature rectangle.</summary>
    public OfficeImageFit ImageFit { get; set; } = OfficeImageFit.Contain;

    /// <summary>Padding between the image and the signature rectangle edges, in PDF points.</summary>
    public double ImagePadding {
        get => _imagePadding;
        set => _imagePadding = ValidateNonNegative(value, nameof(value));
    }

    internal byte[]? GetImageBytes() => _imageBytes is null ? null : (byte[])_imageBytes.Clone();

    private static double ValidatePositive(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) {
            throw new ArgumentOutOfRangeException(parameterName, "Value must be finite and positive.");
        }

        return value;
    }

    private static double ValidateNonNegative(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0) {
            throw new ArgumentOutOfRangeException(parameterName, "Value must be finite and non-negative.");
        }

        return value;
    }
}

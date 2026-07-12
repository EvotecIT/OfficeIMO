namespace OfficeIMO.Pdf;

/// <summary>Visible signature widget and dependency-free appearance-stream settings.</summary>
public sealed class PdfVisibleSignatureAppearanceOptions {
    private int _pageNumber = 1;
    private double _width = 180;
    private double _height = 48;
    private double _fontSize = 10;

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

    private static double ValidatePositive(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) {
            throw new ArgumentOutOfRangeException(parameterName, "Value must be finite and positive.");
        }

        return value;
    }
}

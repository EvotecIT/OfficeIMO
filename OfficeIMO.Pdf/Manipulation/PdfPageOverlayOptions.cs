namespace OfficeIMO.Pdf;

/// <summary>Controls placement of a source PDF page imported as a Form XObject onto target pages.</summary>
public sealed class PdfPageOverlayOptions {
    private int _sourcePageNumber = 1;
    private double? _x;
    private double? _y;
    private double? _width;
    private double? _height;
    private double _opacity = 1D;
    private PdfAlign _horizontalAlignment = PdfAlign.Center;
    private PdfVerticalAlign _verticalAlignment = PdfVerticalAlign.Middle;
    private PdfPageOverlayFit _fit = PdfPageOverlayFit.Contain;

    /// <summary>One-based page number imported from the source PDF.</summary>
    public int SourcePageNumber {
        get => _sourcePageNumber;
        set {
            if (value < 1) throw new ArgumentOutOfRangeException(nameof(SourcePageNumber), value, "Source page number must be 1 or greater.");
            _sourcePageNumber = value;
        }
    }

    /// <summary>Optional target-page selector. Null applies the imported page to every target page.</summary>
    public PdfPageSelector? TargetPages { get; set; }
    /// <summary>How the imported page fits the target rectangle.</summary>
    public PdfPageOverlayFit Fit {
        get => _fit;
        set {
            if (value != PdfPageOverlayFit.None && value != PdfPageOverlayFit.Contain && value != PdfPageOverlayFit.Cover && value != PdfPageOverlayFit.Stretch) {
                throw new ArgumentOutOfRangeException(nameof(Fit), value, "Unsupported imported-page fit mode.");
            }
            _fit = value;
        }
    }
    /// <summary>Horizontal alignment inside the target page or target rectangle.</summary>
    public PdfAlign HorizontalAlignment {
        get => _horizontalAlignment;
        set {
            Guard.LeftCenterRightAlign(value, nameof(HorizontalAlignment), "Imported PDF page");
            _horizontalAlignment = value;
        }
    }
    /// <summary>Vertical alignment inside the target page or target rectangle.</summary>
    public PdfVerticalAlign VerticalAlignment {
        get => _verticalAlignment;
        set {
            if (value != PdfVerticalAlign.Top && value != PdfVerticalAlign.Middle && value != PdfVerticalAlign.Bottom) throw new ArgumentOutOfRangeException(nameof(VerticalAlignment), value, "Unsupported imported-page vertical alignment.");
            _verticalAlignment = value;
        }
    }
    /// <summary>Optional target rectangle X coordinate in points.</summary>
    public double? X { get => _x; set { ValidateOptionalFinite(value, nameof(X)); _x = value; } }
    /// <summary>Optional target rectangle Y coordinate in points.</summary>
    public double? Y { get => _y; set { ValidateOptionalFinite(value, nameof(Y)); _y = value; } }
    /// <summary>Optional target rectangle width in points.</summary>
    public double? Width { get => _width; set { ValidateOptionalPositive(value, nameof(Width)); _width = value; } }
    /// <summary>Optional target rectangle height in points.</summary>
    public double? Height { get => _height; set { ValidateOptionalPositive(value, nameof(Height)); _height = value; } }
    /// <summary>Imported page opacity from 0 through 1.</summary>
    public double Opacity {
        get => _opacity;
        set {
            if (value < 0D || value > 1D || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(nameof(Opacity), value, "Imported page opacity must be finite and between 0 and 1.");
            _opacity = value;
        }
    }
    /// <summary>Places the imported page before existing page content.</summary>
    public bool BehindContent { get; set; }

    /// <summary>Read options used for the imported source PDF, including its password and permission policy.</summary>
    public PdfReadOptions? SourceReadOptions { get; set; }

    /// <summary>Sets the target pages from a rich page-selector expression.</summary>
    public PdfPageOverlayOptions UseTargetPages(string selector) {
        TargetPages = PdfPageSelector.Parse(selector);
        return this;
    }

    internal PdfPageOverlayOptions Clone(bool? behindContent = null) => new PdfPageOverlayOptions {
        SourcePageNumber = SourcePageNumber,
        TargetPages = TargetPages,
        Fit = Fit,
        HorizontalAlignment = HorizontalAlignment,
        VerticalAlignment = VerticalAlignment,
        X = X,
        Y = Y,
        Width = Width,
        Height = Height,
        Opacity = Opacity,
        BehindContent = behindContent ?? BehindContent,
        SourceReadOptions = SourceReadOptions
    };

    private static void ValidateOptionalFinite(double? value, string paramName) {
        if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value))) throw new ArgumentOutOfRangeException(paramName, "Imported-page coordinates must be finite.");
    }

    private static void ValidateOptionalPositive(double? value, string paramName) {
        if (value.HasValue && (value.Value <= 0D || double.IsNaN(value.Value) || double.IsInfinity(value.Value))) throw new ArgumentOutOfRangeException(paramName, "Imported-page dimensions must be positive and finite.");
    }
}

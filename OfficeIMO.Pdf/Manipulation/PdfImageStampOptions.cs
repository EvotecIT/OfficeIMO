namespace OfficeIMO.Pdf;

/// <summary>
/// Options used when adding an image stamp or image watermark to parsed PDF pages.
/// </summary>
public sealed class PdfImageStampOptions {
    private int[]? _pageNumbers;
    private double? _x;
    private double? _y;
    private double? _width;
    private double? _height;
    private double _rotationDegrees;

    /// <summary>
    /// One-based page numbers to stamp. When null or empty, every page is stamped.
    /// </summary>
    public int[]? PageNumbers {
        get => _pageNumbers is null ? null : (int[])_pageNumbers.Clone();
        set {
            ValidatePageNumbers(value, nameof(PageNumbers));
            _pageNumbers = value is null ? null : (int[])value.Clone();
        }
    }

    /// <summary>
    /// Selects an inclusive one-based page range to stamp.
    /// </summary>
    public PdfImageStampOptions UsePageRange(int firstPage, int lastPage) {
        PageNumbers = PdfStampPageSelection.BuildInclusivePageRange(firstPage, lastPage);
        return this;
    }

    /// <summary>
    /// Selects an inclusive one-based page range to stamp.
    /// </summary>
    public PdfImageStampOptions UsePageRange(PdfPageRange pageRange) {
        PageNumbers = PdfStampPageSelection.BuildInclusivePageRange(pageRange);
        return this;
    }

    /// <summary>
    /// Selects inclusive one-based page ranges to stamp.
    /// Overlapping ranges are treated as one page selection set in first-seen order.
    /// </summary>
    public PdfImageStampOptions UsePageRanges(params PdfPageRange[] pageRanges) {
        PageNumbers = PdfStampPageSelection.BuildInclusivePageRanges(pageRanges);
        return this;
    }

    /// <summary>
    /// X coordinate of the image origin in PDF points. When null, a sensible default is used.
    /// </summary>
    public double? X {
        get => _x;
        set {
            ValidateOptionalFinite(value, nameof(X), "Image stamp X coordinate must be finite.");
            _x = value;
        }
    }

    /// <summary>
    /// Y coordinate of the image origin in PDF points. When null, a sensible default is used.
    /// </summary>
    public double? Y {
        get => _y;
        set {
            ValidateOptionalFinite(value, nameof(Y), "Image stamp Y coordinate must be finite.");
            _y = value;
        }
    }

    /// <summary>
    /// Width of the stamped image in PDF points. When null, the image pixel width is used.
    /// </summary>
    public double? Width {
        get => _width;
        set {
            ValidateOptionalPositiveFinite(value, nameof(Width), "Image stamp width must be a positive finite value.");
            _width = value;
        }
    }

    /// <summary>
    /// Height of the stamped image in PDF points. When null, the image pixel height is used.
    /// </summary>
    public double? Height {
        get => _height;
        set {
            ValidateOptionalPositiveFinite(value, nameof(Height), "Image stamp height must be a positive finite value.");
            _height = value;
        }
    }

    /// <summary>
    /// Image rotation in degrees around the image origin.
    /// </summary>
    public double RotationDegrees {
        get => _rotationDegrees;
        set {
            ValidateFinite(value, nameof(RotationDegrees), "Image stamp rotation must be finite.");
            _rotationDegrees = value;
        }
    }

    /// <summary>
    /// Places the new image content stream before existing page content when true.
    /// </summary>
    public bool BehindContent { get; set; }

    private static void ValidateOptionalFinite(double? value, string paramName, string message) {
        if (value.HasValue) {
            ValidateFinite(value.Value, paramName, message);
        }
    }

    private static void ValidateOptionalPositiveFinite(double? value, string paramName, string message) {
        if (value.HasValue && (value.Value <= 0 || double.IsNaN(value.Value) || double.IsInfinity(value.Value))) {
            throw new System.ArgumentOutOfRangeException(paramName, message);
        }
    }

    private static void ValidateFinite(double value, string paramName, string message) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentOutOfRangeException(paramName, message);
        }
    }

    private static void ValidatePageNumbers(int[]? pageNumbers, string paramName) {
        if (pageNumbers is null) {
            return;
        }

        var seen = new System.Collections.Generic.HashSet<int>();
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            if (pageNumber < 1) {
                throw new System.ArgumentOutOfRangeException(paramName, "Stamp page numbers must be 1 or greater.");
            }

            if (!seen.Add(pageNumber)) {
                throw new System.ArgumentException("Duplicate page selections are not supported.", paramName);
            }
        }
    }
}

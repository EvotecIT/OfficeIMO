namespace OfficeIMO.Pdf;

/// <summary>
/// Options used when adding a simple text stamp or watermark to parsed PDF pages.
/// </summary>
public sealed class PdfTextStampOptions {
    private int[]? _pageNumbers;
    private double? _x;
    private double? _y;
    private PdfStandardFont _font = PdfStandardFont.HelveticaBold;
    private double _fontSize = 24;
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
    public PdfTextStampOptions UsePageRange(int firstPage, int lastPage) {
        PageNumbers = PdfStampPageSelection.BuildInclusivePageRange(firstPage, lastPage);
        return this;
    }

    /// <summary>
    /// Selects an inclusive one-based page range to stamp.
    /// </summary>
    public PdfTextStampOptions UsePageRange(PdfPageRange pageRange) {
        PageNumbers = PdfStampPageSelection.BuildInclusivePageRange(pageRange);
        return this;
    }

    /// <summary>
    /// Selects inclusive one-based page ranges to stamp.
    /// Overlapping ranges are treated as one page selection set in first-seen order.
    /// </summary>
    public PdfTextStampOptions UsePageRanges(params PdfPageRange[] pageRanges) {
        PageNumbers = PdfStampPageSelection.BuildInclusivePageRanges(pageRanges);
        return this;
    }

    /// <summary>
    /// X coordinate of the text baseline origin in PDF points. When null, a sensible default is used.
    /// </summary>
    public double? X {
        get => _x;
        set {
            ValidateOptionalFinite(value, nameof(X), "Text stamp X coordinate must be finite.");
            _x = value;
        }
    }

    /// <summary>
    /// Y coordinate of the text baseline origin in PDF points. When null, a sensible default is used.
    /// </summary>
    public double? Y {
        get => _y;
        set {
            ValidateOptionalFinite(value, nameof(Y), "Text stamp Y coordinate must be finite.");
            _y = value;
        }
    }

    /// <summary>
    /// Standard PDF font used for the stamp.
    /// </summary>
    public PdfStandardFont Font {
        get => _font;
        set {
            Guard.StandardFont(value, nameof(Font), "Text stamp font must be one of the supported standard PDF fonts.");
            _font = value;
        }
    }

    /// <summary>
    /// Text size in PDF points.
    /// </summary>
    public double FontSize {
        get => _fontSize;
        set {
            ValidatePositiveFinite(value, nameof(FontSize), "Font size must be a positive finite value.");
            _fontSize = value;
        }
    }

    /// <summary>
    /// RGB text color.
    /// </summary>
    public PdfColor Color { get; set; } = PdfColor.Gray;

    /// <summary>
    /// Text rotation in degrees around the text origin.
    /// </summary>
    public double RotationDegrees {
        get => _rotationDegrees;
        set {
            ValidateFinite(value, nameof(RotationDegrees), "Text stamp rotation must be finite.");
            _rotationDegrees = value;
        }
    }

    /// <summary>
    /// Places the new content stream before existing page content when true.
    /// </summary>
    public bool BehindContent { get; set; }

    private static void ValidateOptionalFinite(double? value, string paramName, string message) {
        if (value.HasValue) {
            ValidateFinite(value.Value, paramName, message);
        }
    }

    private static void ValidatePositiveFinite(double value, string paramName, string message) {
        if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) {
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

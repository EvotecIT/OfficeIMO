namespace OfficeIMO.Pdf;

/// <summary>
/// Options used when adding an image stamp or image watermark to parsed PDF pages.
/// </summary>
public sealed class PdfImageStampOptions {
    /// <summary>Cancels input reading, PDF processing, and output publication for this operation.</summary>
    public System.Threading.CancellationToken CancellationToken { get; set; }

    private int[]? _pageNumbers;
    private global::OfficeIMO.OfficeLength? _width;
    private global::OfficeIMO.OfficeLength? _height;
    private double _rotationDegrees;

    /// <summary>Maximum encoded image bytes accepted for a stamp or watermark. Defaults to 128 MiB.</summary>
    public long MaximumEncodedImageBytes { get; set; } = PdfImageInput.DefaultMaximumEncodedBytes;

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
    /// Physical X coordinate of the image origin from the unrotated PDF origin.
    /// When null, stamps use 36 points and watermarks center in the effective page box.
    /// </summary>
    public global::OfficeIMO.OfficeLength? X { get; set; }

    /// <summary>
    /// Physical Y coordinate of the image origin from the unrotated PDF origin.
    /// When null, stamps use 36 points and watermarks center in the effective page box.
    /// </summary>
    public global::OfficeIMO.OfficeLength? Y { get; set; }

    /// <summary>
    /// Physical width of the stamped image. When null, each image pixel contributes one point.
    /// </summary>
    public global::OfficeIMO.OfficeLength? Width {
        get => _width;
        set {
            ValidateOptionalPositive(value, nameof(Width), "Image stamp width must be a positive finite value.");
            _width = value;
        }
    }

    /// <summary>
    /// Physical height of the stamped image. When null, each image pixel contributes one point.
    /// </summary>
    public global::OfficeIMO.OfficeLength? Height {
        get => _height;
        set {
            ValidateOptionalPositive(value, nameof(Height), "Image stamp height must be a positive finite value.");
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

    private static void ValidateOptionalPositive(global::OfficeIMO.OfficeLength? value, string paramName, string message) {
        if (value.HasValue && value.Value.ToEmus() <= 0) {
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

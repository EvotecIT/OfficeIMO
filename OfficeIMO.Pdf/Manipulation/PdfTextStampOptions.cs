namespace OfficeIMO.Pdf;

/// <summary>
/// Options used when adding a simple text stamp or watermark to parsed PDF pages.
/// </summary>
public sealed class PdfTextStampOptions {
    /// <summary>Cancels input reading, PDF processing, and output publication for this operation.</summary>
    public System.Threading.CancellationToken CancellationToken { get; set; }

    private int[]? _pageNumbers;
    private PdfAlign _align = PdfAlign.Left;
    private PdfStandardFont _font = PdfStandardFont.HelveticaBold;
    private global::OfficeIMO.OfficeLength _fontSize = global::OfficeIMO.OfficeLength.Points(24);
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
    /// Physical X coordinate of the text baseline origin from the unrotated PDF origin.
    /// When null, stamps use 36 points and watermarks center in the effective page box.
    /// </summary>
    public global::OfficeIMO.OfficeLength? X { get; set; }

    /// <summary>
    /// Physical Y coordinate of the text baseline origin from the unrotated PDF origin.
    /// When null, stamps use 36 points and watermarks center in the effective page box.
    /// </summary>
    public global::OfficeIMO.OfficeLength? Y { get; set; }

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
    /// Physical text size. Defaults to 24 points.
    /// </summary>
    public global::OfficeIMO.OfficeLength FontSize {
        get => _fontSize;
        set {
            ValidatePositive(value, nameof(FontSize), "Font size must be a positive finite value.");
            _fontSize = value;
        }
    }

    /// <summary>
    /// RGB text color.
    /// </summary>
    public global::OfficeIMO.OfficeColor Color { get; set; } = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.5D, 0.5D, 0.5D);

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

    /// <summary>
    /// Horizontal alignment of the text against <see cref="X"/>: <see cref="PdfAlign.Left"/> (the default) starts the
    /// text at X, <see cref="PdfAlign.Center"/> centres it on X and <see cref="PdfAlign.Right"/> ends it at X. When X is
    /// null, centred text uses the page centre and right-aligned text ends 36 points from the right page edge.
    /// </summary>
    public PdfAlign Align {
        get => _align;
        set {
            Guard.LeftCenterRightAlign(value, nameof(Align), "Text stamp");
            _align = value;
        }
    }

    /// <summary>
    /// When true (the default), the stamp text uses the installed font that backs <see cref="Font"/> in the default
    /// theme, the same font documents created by OfficeIMO.Pdf use, embedded as a subset. When no such font is
    /// installed, or it lacks a character of the text, the non-embedded standard font is written instead.
    /// Set to false to always write the non-embedded standard font.
    /// </summary>
    public bool EmbedFont { get; set; } = PdfCreateOptions.EmbedDefaultFontsInitially;

    private static void ValidatePositive(global::OfficeIMO.OfficeLength value, string paramName, string message) {
        if (value.ToEmus() <= 0) {
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

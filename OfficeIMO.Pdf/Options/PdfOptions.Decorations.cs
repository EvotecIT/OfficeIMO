namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    /// <summary>Optional page background color rendered behind all page content.</summary>
    public PdfColor? BackgroundColor { get; set; }
    /// <summary>Optional reusable text watermark rendered behind all page content.</summary>
    public PdfTextWatermark? TextWatermark {
        get => _textWatermark?.Clone();
        set => _textWatermark = value?.Clone();
    }
    internal PdfTextWatermark? TextWatermarkSnapshot => _textWatermark?.Clone();
    /// <summary>Optional first-page text watermark rendered behind page content when first-page variants are enabled.</summary>
    public PdfTextWatermark? FirstPageTextWatermark {
        get => _firstPageTextWatermark?.Clone();
        set {
            _firstPageTextWatermark = value?.Clone();
            _suppressFirstPageTextWatermark = false;
        }
    }
    /// <summary>Optional even-page text watermark rendered behind page content when odd/even variants are enabled.</summary>
    public PdfTextWatermark? EvenPageTextWatermark {
        get => _evenPageTextWatermark?.Clone();
        set {
            _evenPageTextWatermark = value?.Clone();
            _suppressEvenPageTextWatermark = false;
        }
    }
    internal PdfTextWatermark? GetTextWatermarkForPage(int pageNumber) {
        if (pageNumber == 1 && (_firstPageTextWatermark != null || _suppressFirstPageTextWatermark)) {
            return _suppressFirstPageTextWatermark
                ? null
                : _firstPageTextWatermark?.Clone();
        }

        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            if (_suppressFirstPageTextWatermark) {
                return null;
            }

            return (_firstPageTextWatermark ?? _textWatermark)?.Clone();
        }

        if (pageNumber > 0 && pageNumber % 2 == 0 && (_evenPageTextWatermark != null || _suppressEvenPageTextWatermark)) {
            return _suppressEvenPageTextWatermark
                ? null
                : _evenPageTextWatermark?.Clone();
        }

        if (DifferentOddAndEvenPagesHeaderFooter && pageNumber > 0 && pageNumber % 2 == 0) {
            if (_suppressEvenPageTextWatermark) {
                return null;
            }

            return (_evenPageTextWatermark ?? _textWatermark)?.Clone();
        }

        return _textWatermark?.Clone();
    }
    /// <summary>Optional reusable image watermark rendered behind all page content.</summary>
    public PdfImageWatermark? ImageWatermark {
        get => _imageWatermark?.Clone();
        set => _imageWatermark = value?.Clone();
    }
    internal PdfImageWatermark? ImageWatermarkSnapshot => _imageWatermark?.Clone();
    /// <summary>Optional first-page image watermark rendered behind page content when first-page variants are enabled.</summary>
    public PdfImageWatermark? FirstPageImageWatermark {
        get => _firstPageImageWatermark?.Clone();
        set {
            _firstPageImageWatermark = value?.Clone();
            _suppressFirstPageImageWatermark = false;
        }
    }
    /// <summary>Optional even-page image watermark rendered behind page content when odd/even variants are enabled.</summary>
    public PdfImageWatermark? EvenPageImageWatermark {
        get => _evenPageImageWatermark?.Clone();
        set {
            _evenPageImageWatermark = value?.Clone();
            _suppressEvenPageImageWatermark = false;
        }
    }
    internal PdfImageWatermark? GetImageWatermarkForPage(int pageNumber) {
        if (pageNumber == 1 && (_firstPageImageWatermark != null || _suppressFirstPageImageWatermark)) {
            return _suppressFirstPageImageWatermark
                ? null
                : _firstPageImageWatermark?.Clone();
        }

        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            if (_suppressFirstPageImageWatermark) {
                return null;
            }

            return (_firstPageImageWatermark ?? _imageWatermark)?.Clone();
        }

        if (pageNumber > 0 && pageNumber % 2 == 0 && (_evenPageImageWatermark != null || _suppressEvenPageImageWatermark)) {
            return _suppressEvenPageImageWatermark
                ? null
                : _evenPageImageWatermark?.Clone();
        }

        if (DifferentOddAndEvenPagesHeaderFooter && pageNumber > 0 && pageNumber % 2 == 0) {
            if (_suppressEvenPageImageWatermark) {
                return null;
            }

            return (_evenPageImageWatermark ?? _imageWatermark)?.Clone();
        }

        return _imageWatermark?.Clone();
    }
    internal void SuppressFirstPageTextWatermark() {
        _firstPageTextWatermark = null;
        _suppressFirstPageTextWatermark = true;
    }

    internal void SuppressEvenPageTextWatermark() {
        _evenPageTextWatermark = null;
        _suppressEvenPageTextWatermark = true;
    }

    internal void SuppressFirstPageImageWatermark() {
        _firstPageImageWatermark = null;
        _suppressFirstPageImageWatermark = true;
    }

    internal void SuppressEvenPageImageWatermark() {
        _evenPageImageWatermark = null;
        _suppressEvenPageImageWatermark = true;
    }
    /// <summary>Optional reusable page border rendered as a page decoration.</summary>
    public PdfPageBorder? PageBorder {
        get => _pageBorder?.Clone();
        set => _pageBorder = value?.Clone();
    }
    internal PdfPageBorder? PageBorderSnapshot => _pageBorder?.Clone();
    /// <summary>Optional reusable page background image rendered behind all page content.</summary>
    public PdfPageBackgroundImage? PageBackgroundImage {
        get => _pageBackgroundImage?.Clone();
        set => _pageBackgroundImage = value?.Clone();
    }
    internal PdfPageBackgroundImage? PageBackgroundImageSnapshot => _pageBackgroundImage?.Clone();
    /// <summary>Optional reusable page background shapes rendered behind all page content.</summary>
    public System.Collections.Generic.IReadOnlyList<PdfPageBackgroundShape>? PageBackgroundShapes {
        get => ClonePageBackgroundShapes(_pageBackgroundShapes);
        set => _pageBackgroundShapes = ClonePageBackgroundShapes(value);
    }
    internal System.Collections.Generic.IReadOnlyList<PdfPageBackgroundShape> PageBackgroundShapeSnapshots =>
        (System.Collections.Generic.IReadOnlyList<PdfPageBackgroundShape>?)ClonePageBackgroundShapes(_pageBackgroundShapes) ?? System.Array.Empty<PdfPageBackgroundShape>();
    internal void AddPageBackgroundShape(PdfPageBackgroundShape shape) {
        Guard.NotNull(shape, nameof(shape));
        (_pageBackgroundShapes ??= new System.Collections.Generic.List<PdfPageBackgroundShape>()).Add(shape.Clone());
    }

    internal void ClearPageBackgroundShapes() {
        _pageBackgroundShapes = null;
    }

}

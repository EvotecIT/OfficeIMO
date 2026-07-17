namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    /// <summary>When true, renders header text using <see cref="HeaderFormat"/>.</summary>
    public bool ShowHeader { get; set; }
    /// <summary>Header text format, supports {page} and {pages}. Default: empty.</summary>
    public string HeaderFormat { get; set; } = string.Empty;
    /// <summary>When true, page 1 uses first-page header/footer content instead of the running header/footer.</summary>
    public bool DifferentFirstPageHeaderFooter { get; set; }
    /// <summary>Header text format used on page 1 when <see cref="DifferentFirstPageHeaderFooter"/> is true. Supports {page} and {pages}.</summary>
    public string FirstPageHeaderFormat { get; set; } = string.Empty;
    /// <summary>When true, even-numbered pages use even-page header/footer content instead of the running odd-page content.</summary>
    public bool DifferentOddAndEvenPagesHeaderFooter { get; set; }
    /// <summary>Header text format used on even-numbered pages when <see cref="DifferentOddAndEvenPagesHeaderFooter"/> is true. Supports {page} and {pages}.</summary>
    public string EvenPageHeaderFormat { get; set; } = string.Empty;
    /// <summary>Header font.</summary>
    public PdfStandardFont HeaderFont {
        get => _headerFont;
        set {
            Guard.StandardFont(value, nameof(HeaderFont), "PDF header font must be one of the supported standard PDF fonts.");
            _headerFont = value;
            _hasExplicitHeaderFont = true;
        }
    }
    /// <summary>Gets whether the header font slot was explicitly supplied by the caller or a theme.</summary>
    public bool HasExplicitHeaderFont => _hasExplicitHeaderFont;
    /// <summary>Header font size in points.</summary>
    public double HeaderFontSize { get; set; } = 9;
    /// <summary>Header text color. When null, the current PDF fill color is preserved.</summary>
    public PdfColor? HeaderTextColor { get; set; }
    /// <summary>Header alignment.</summary>
    public PdfAlign HeaderAlign {
        get => _headerAlign;
        set {
            Guard.LeftCenterRightAlign(value, nameof(HeaderAlign), "PDF header");
            _headerAlign = value;
        }
    }
    /// <summary>Header baseline Y offset above the top margin (points). Default 18.</summary>
    public double HeaderOffsetY { get; set; } = 18;

    /// <summary>When true, renders page numbers in the footer using <see cref="FooterFormat"/>.</summary>
    public bool ShowPageNumbers { get; set; } // default false
    /// <summary>Footer text format, supports {page} and {pages}. Default: "Page {page}/{pages}".</summary>
    public string FooterFormat { get; set; } = "Page {page}/{pages}";
    /// <summary>Footer text format used on page 1 when <see cref="DifferentFirstPageHeaderFooter"/> is true. Supports {page} and {pages}.</summary>
    public string FirstPageFooterFormat { get; set; } = string.Empty;
    /// <summary>Footer text format used on even-numbered pages when <see cref="DifferentOddAndEvenPagesHeaderFooter"/> is true. Supports {page} and {pages}.</summary>
    public string EvenPageFooterFormat { get; set; } = string.Empty;
    /// <summary>Footer font.</summary>
    public PdfStandardFont FooterFont {
        get => _footerFont;
        set {
            Guard.StandardFont(value, nameof(FooterFont), "PDF footer font must be one of the supported standard PDF fonts.");
            _footerFont = value;
            _hasExplicitFooterFont = true;
        }
    }
    /// <summary>Gets whether the footer font slot was explicitly supplied by the caller or a theme.</summary>
    public bool HasExplicitFooterFont => _hasExplicitFooterFont;
    /// <summary>Footer font size in points.</summary>
    public double FooterFontSize { get; set; } = 9;
    /// <summary>Footer text color. When null, the current PDF fill color is preserved.</summary>
    public PdfColor? FooterTextColor { get; set; }
    /// <summary>Footer alignment.</summary>
    public PdfAlign FooterAlign {
        get => _footerAlign;
        set {
            Guard.LeftCenterRightAlign(value, nameof(FooterAlign), "PDF footer");
            _footerAlign = value;
        }
    }
    /// <summary>Footer baseline Y position from bottom margin (points). Default 18.</summary>
    public double FooterOffsetY { get; set; } = 18;

    /// <summary>First visible page number for this document or section flow. Default 1.</summary>
    public int PageNumberStart {
        get => _pageNumberStart;
        set {
            if (value < 1) {
                throw new System.ArgumentOutOfRangeException(nameof(PageNumberStart), "PDF page number start must be a positive value.");
            }

            _pageNumberStart = value;
            _hasExplicitPageNumberStart = true;
        }
    }
    internal bool HasExplicitPageNumberStart => _hasExplicitPageNumberStart;

    /// <summary>Visible numbering style for generated page tokens. Default Arabic.</summary>
    public PdfPageNumberStyle PageNumberStyle {
        get => _pageNumberStyle;
        set {
            Guard.PageNumberStyle(value, nameof(PageNumberStyle));
            _pageNumberStyle = value;
        }
    }

    /// <summary>Advanced footer template segments. When set, overrides FooterFormat.</summary>
    public System.Collections.Generic.List<FooterSegment>? FooterSegments {
        get => _footerSegments is null ? null : new System.Collections.Generic.List<FooterSegment>(_footerSegments);
        set => _footerSegments = value is null ? null : new System.Collections.Generic.List<FooterSegment>(value);
    }

    /// <summary>Advanced page-1 footer template segments used when <see cref="DifferentFirstPageHeaderFooter"/> is true.</summary>
    public System.Collections.Generic.List<FooterSegment>? FirstPageFooterSegments {
        get => _firstPageFooterSegments is null ? null : new System.Collections.Generic.List<FooterSegment>(_firstPageFooterSegments);
        set => _firstPageFooterSegments = value is null ? null : new System.Collections.Generic.List<FooterSegment>(value);
    }

    /// <summary>Advanced even-page footer template segments used when <see cref="DifferentOddAndEvenPagesHeaderFooter"/> is true.</summary>
    public System.Collections.Generic.List<FooterSegment>? EvenPageFooterSegments {
        get => _evenPageFooterSegments is null ? null : new System.Collections.Generic.List<FooterSegment>(_evenPageFooterSegments);
        set => _evenPageFooterSegments = value is null ? null : new System.Collections.Generic.List<FooterSegment>(value);
    }
    internal bool HasHeaderContent => (ShowHeader && HeaderFormat != null && HeaderFormat.Length > 0) ||
        (_headerSegments != null && _headerSegments.Count > 0) ||
        HasHeaderZoneContent ||
        HasHeaderImageContent;
    internal bool HasFooterContent => ShowPageNumbers ||
        (_footerSegments != null && _footerSegments.Count > 0) ||
        HasFooterZoneContent ||
        HasFooterImageContent;
    internal bool HasHeaderZoneContent =>
        !string.IsNullOrEmpty(_headerLeftFormat) ||
        !string.IsNullOrEmpty(_headerCenterFormat) ||
        !string.IsNullOrEmpty(_headerRightFormat);
    internal bool HasFirstPageHeaderZoneContent =>
        !string.IsNullOrEmpty(_firstPageHeaderLeftFormat) ||
        !string.IsNullOrEmpty(_firstPageHeaderCenterFormat) ||
        !string.IsNullOrEmpty(_firstPageHeaderRightFormat);
    internal bool HasEvenPageHeaderZoneContent =>
        !string.IsNullOrEmpty(_evenPageHeaderLeftFormat) ||
        !string.IsNullOrEmpty(_evenPageHeaderCenterFormat) ||
        !string.IsNullOrEmpty(_evenPageHeaderRightFormat);
    internal bool HasFooterZoneContent =>
        !string.IsNullOrEmpty(_footerLeftFormat) ||
        !string.IsNullOrEmpty(_footerCenterFormat) ||
        !string.IsNullOrEmpty(_footerRightFormat);
    internal bool HasFirstPageFooterZoneContent =>
        !string.IsNullOrEmpty(_firstPageFooterLeftFormat) ||
        !string.IsNullOrEmpty(_firstPageFooterCenterFormat) ||
        !string.IsNullOrEmpty(_firstPageFooterRightFormat);
    internal bool HasEvenPageFooterZoneContent =>
        !string.IsNullOrEmpty(_evenPageFooterLeftFormat) ||
        !string.IsNullOrEmpty(_evenPageFooterCenterFormat) ||
        !string.IsNullOrEmpty(_evenPageFooterRightFormat);
    internal bool HasHeaderImageContent => _headerImages != null && _headerImages.Count > 0;
    internal bool HasFirstPageHeaderImageContent => _firstPageHeaderImages != null && _firstPageHeaderImages.Count > 0;
    internal bool HasEvenPageHeaderImageContent => _evenPageHeaderImages != null && _evenPageHeaderImages.Count > 0;
    internal bool HasFooterImageContent => _footerImages != null && _footerImages.Count > 0;
    internal bool HasFirstPageFooterImageContent => _firstPageFooterImages != null && _firstPageFooterImages.Count > 0;
    internal bool HasEvenPageFooterImageContent => _evenPageFooterImages != null && _evenPageFooterImages.Count > 0;
    internal bool HasHeaderContentForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return (FirstPageHeaderFormat != null && FirstPageHeaderFormat.Length > 0) ||
                (_firstPageHeaderSegments != null && _firstPageHeaderSegments.Count > 0) ||
                HasFirstPageHeaderZoneContent ||
                HasFirstPageHeaderImageContent;
        }

        if (IsEvenPageVariant(pageNumber)) {
            return (EvenPageHeaderFormat != null && EvenPageHeaderFormat.Length > 0) ||
                (_evenPageHeaderSegments != null && _evenPageHeaderSegments.Count > 0) ||
                HasEvenPageHeaderZoneContent ||
                HasEvenPageHeaderImageContent;
        }

        return HasHeaderContent;
    }

    internal bool HasHeaderTextContentForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return (FirstPageHeaderFormat != null && FirstPageHeaderFormat.Length > 0) ||
                (_firstPageHeaderSegments != null && _firstPageHeaderSegments.Count > 0) ||
                HasFirstPageHeaderZoneContent;
        }

        if (IsEvenPageVariant(pageNumber)) {
            return (EvenPageHeaderFormat != null && EvenPageHeaderFormat.Length > 0) ||
                (_evenPageHeaderSegments != null && _evenPageHeaderSegments.Count > 0) ||
                HasEvenPageHeaderZoneContent;
        }

        return (ShowHeader && HeaderFormat != null && HeaderFormat.Length > 0) ||
            (_headerSegments != null && _headerSegments.Count > 0) ||
            HasHeaderZoneContent;
    }

    internal bool HasFooterContentForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return (FirstPageFooterFormat != null && FirstPageFooterFormat.Length > 0) ||
                (_firstPageFooterSegments != null && _firstPageFooterSegments.Count > 0) ||
                HasFirstPageFooterZoneContent ||
                HasFirstPageFooterImageContent;
        }

        if (IsEvenPageVariant(pageNumber)) {
            return (EvenPageFooterFormat != null && EvenPageFooterFormat.Length > 0) ||
                (_evenPageFooterSegments != null && _evenPageFooterSegments.Count > 0) ||
                HasEvenPageFooterZoneContent ||
                HasEvenPageFooterImageContent;
        }

        return HasFooterContent;
    }

    internal bool HasFooterTextContentForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return (FirstPageFooterFormat != null && FirstPageFooterFormat.Length > 0) ||
                (_firstPageFooterSegments != null && _firstPageFooterSegments.Count > 0) ||
                HasFirstPageFooterZoneContent;
        }

        if (IsEvenPageVariant(pageNumber)) {
            return (EvenPageFooterFormat != null && EvenPageFooterFormat.Length > 0) ||
                (_evenPageFooterSegments != null && _evenPageFooterSegments.Count > 0) ||
                HasEvenPageFooterZoneContent;
        }

        return ShowPageNumbers ||
            (_footerSegments != null && _footerSegments.Count > 0) ||
            HasFooterZoneContent;
    }

    internal string GetHeaderFormatForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return FirstPageHeaderFormat;
        }

        return IsEvenPageVariant(pageNumber) ? EvenPageHeaderFormat : HeaderFormat;
    }

    internal System.Collections.Generic.IReadOnlyList<FooterSegment>? GetHeaderSegmentsForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return _firstPageHeaderSegments;
        }

        return IsEvenPageVariant(pageNumber) ? _evenPageHeaderSegments : _headerSegments;
    }

    internal (string? Left, string? Center, string? Right) GetHeaderZonesForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return (_firstPageHeaderLeftFormat, _firstPageHeaderCenterFormat, _firstPageHeaderRightFormat);
        }

        if (IsEvenPageVariant(pageNumber)) {
            return (_evenPageHeaderLeftFormat, _evenPageHeaderCenterFormat, _evenPageHeaderRightFormat);
        }

        return (_headerLeftFormat, _headerCenterFormat, _headerRightFormat);
    }

    internal System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage> GetHeaderImagesForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return _firstPageHeaderImages != null ? _firstPageHeaderImages : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage>)System.Array.Empty<PdfHeaderFooterImage>();
        }

        if (IsEvenPageVariant(pageNumber)) {
            return _evenPageHeaderImages != null ? _evenPageHeaderImages : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage>)System.Array.Empty<PdfHeaderFooterImage>();
        }

        return _headerImages != null ? _headerImages : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage>)System.Array.Empty<PdfHeaderFooterImage>();
    }

    internal System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape> GetHeaderShapesForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return _firstPageHeaderShapes != null ? _firstPageHeaderShapes : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape>)System.Array.Empty<PdfHeaderFooterShape>();
        }

        if (IsEvenPageVariant(pageNumber)) {
            return _evenPageHeaderShapes != null ? _evenPageHeaderShapes : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape>)System.Array.Empty<PdfHeaderFooterShape>();
        }

        return _headerShapes != null ? _headerShapes : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape>)System.Array.Empty<PdfHeaderFooterShape>();
    }

    internal string GetFooterFormatForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return FirstPageFooterFormat;
        }

        return IsEvenPageVariant(pageNumber) ? EvenPageFooterFormat : FooterFormat;
    }

    internal System.Collections.Generic.IReadOnlyList<FooterSegment>? GetFooterSegmentsForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return _firstPageFooterSegments;
        }

        return IsEvenPageVariant(pageNumber) ? _evenPageFooterSegments : _footerSegments;
    }

    internal (string? Left, string? Center, string? Right) GetFooterZonesForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return (_firstPageFooterLeftFormat, _firstPageFooterCenterFormat, _firstPageFooterRightFormat);
        }

        if (IsEvenPageVariant(pageNumber)) {
            return (_evenPageFooterLeftFormat, _evenPageFooterCenterFormat, _evenPageFooterRightFormat);
        }

        return (_footerLeftFormat, _footerCenterFormat, _footerRightFormat);
    }

    internal System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage> GetFooterImagesForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return _firstPageFooterImages != null ? _firstPageFooterImages : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage>)System.Array.Empty<PdfHeaderFooterImage>();
        }

        if (IsEvenPageVariant(pageNumber)) {
            return _evenPageFooterImages != null ? _evenPageFooterImages : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage>)System.Array.Empty<PdfHeaderFooterImage>();
        }

        return _footerImages != null ? _footerImages : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterImage>)System.Array.Empty<PdfHeaderFooterImage>();
    }

    internal System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape> GetFooterShapesForPage(int pageNumber) {
        if (pageNumber == 1 && DifferentFirstPageHeaderFooter) {
            return _firstPageFooterShapes != null ? _firstPageFooterShapes : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape>)System.Array.Empty<PdfHeaderFooterShape>();
        }

        if (IsEvenPageVariant(pageNumber)) {
            return _evenPageFooterShapes != null ? _evenPageFooterShapes : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape>)System.Array.Empty<PdfHeaderFooterShape>();
        }

        return _footerShapes != null ? _footerShapes : (System.Collections.Generic.IReadOnlyList<PdfHeaderFooterShape>)System.Array.Empty<PdfHeaderFooterShape>();
    }

    private bool IsEvenPageVariant(int pageNumber) =>
        DifferentOddAndEvenPagesHeaderFooter && pageNumber > 0 && pageNumber % 2 == 0;

    internal void ClearPageNumberStartOverride() {
        _hasExplicitPageNumberStart = false;
    }

    internal System.Collections.Generic.List<FooterSegment> ResetHeaderSegmentsForCompose() {
        _headerSegments = new System.Collections.Generic.List<FooterSegment>();
        ShowHeader = true;
        return _headerSegments;
    }

    internal void ClearHeaderSegmentsForCompose() {
        _headerSegments = null;
    }

    internal void SetHeaderZonesForCompose(string? left, string? center, string? right) {
        ValidateZones(left, center, right, nameof(left));
        ClearHeaderSegmentsForCompose();
        HeaderFormat = string.Empty;
        ShowHeader = true;
        _headerLeftFormat = left;
        _headerCenterFormat = center;
        _headerRightFormat = right;
    }

    internal void ClearHeaderZonesForCompose() {
        _headerLeftFormat = null;
        _headerCenterFormat = null;
        _headerRightFormat = null;
    }

    internal void AddHeaderImageForCompose(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        ShowHeader = true;
        (_headerImages ??= new System.Collections.Generic.List<PdfHeaderFooterImage>()).Add(image.Clone());
    }

    internal void SetFirstPageHeaderZonesForCompose(string? left, string? center, string? right) {
        ValidateZones(left, center, right, nameof(left));
        ClearFirstPageHeaderSegmentsForCompose();
        DifferentFirstPageHeaderFooter = true;
        FirstPageHeaderFormat = string.Empty;
        _firstPageHeaderLeftFormat = left;
        _firstPageHeaderCenterFormat = center;
        _firstPageHeaderRightFormat = right;
    }

    internal void ClearFirstPageHeaderZonesForCompose() {
        _firstPageHeaderLeftFormat = null;
        _firstPageHeaderCenterFormat = null;
        _firstPageHeaderRightFormat = null;
    }

    internal void AddFirstPageHeaderImageForCompose(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        DifferentFirstPageHeaderFooter = true;
        (_firstPageHeaderImages ??= new System.Collections.Generic.List<PdfHeaderFooterImage>()).Add(image.Clone());
    }

    internal void SetEvenPageHeaderZonesForCompose(string? left, string? center, string? right) {
        ValidateZones(left, center, right, nameof(left));
        ClearEvenPageHeaderSegmentsForCompose();
        DifferentOddAndEvenPagesHeaderFooter = true;
        EvenPageHeaderFormat = string.Empty;
        _evenPageHeaderLeftFormat = left;
        _evenPageHeaderCenterFormat = center;
        _evenPageHeaderRightFormat = right;
    }

    internal void ClearEvenPageHeaderZonesForCompose() {
        _evenPageHeaderLeftFormat = null;
        _evenPageHeaderCenterFormat = null;
        _evenPageHeaderRightFormat = null;
    }

    internal void AddEvenPageHeaderImageForCompose(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        DifferentOddAndEvenPagesHeaderFooter = true;
        (_evenPageHeaderImages ??= new System.Collections.Generic.List<PdfHeaderFooterImage>()).Add(image.Clone());
    }

    internal void AddHeaderShapeForCompose(PdfHeaderFooterShape shape) {
        Guard.NotNull(shape, nameof(shape));
        ShowHeader = true;
        (_headerShapes ??= new System.Collections.Generic.List<PdfHeaderFooterShape>()).Add(shape.Clone());
    }

    internal void AddFirstPageHeaderShapeForCompose(PdfHeaderFooterShape shape) {
        Guard.NotNull(shape, nameof(shape));
        ShowHeader = true;
        DifferentFirstPageHeaderFooter = true;
        (_firstPageHeaderShapes ??= new System.Collections.Generic.List<PdfHeaderFooterShape>()).Add(shape.Clone());
    }

    internal void AddEvenPageHeaderShapeForCompose(PdfHeaderFooterShape shape) {
        Guard.NotNull(shape, nameof(shape));
        ShowHeader = true;
        DifferentOddAndEvenPagesHeaderFooter = true;
        (_evenPageHeaderShapes ??= new System.Collections.Generic.List<PdfHeaderFooterShape>()).Add(shape.Clone());
    }

    internal System.Collections.Generic.List<FooterSegment> ResetFirstPageHeaderSegmentsForCompose() {
        _firstPageHeaderSegments = new System.Collections.Generic.List<FooterSegment>();
        DifferentFirstPageHeaderFooter = true;
        return _firstPageHeaderSegments;
    }

    internal void ClearFirstPageHeaderSegmentsForCompose() {
        _firstPageHeaderSegments = null;
    }

    internal System.Collections.Generic.List<FooterSegment> ResetEvenPageHeaderSegmentsForCompose() {
        _evenPageHeaderSegments = new System.Collections.Generic.List<FooterSegment>();
        DifferentOddAndEvenPagesHeaderFooter = true;
        return _evenPageHeaderSegments;
    }

    internal void ClearEvenPageHeaderSegmentsForCompose() {
        _evenPageHeaderSegments = null;
    }

    internal System.Collections.Generic.List<FooterSegment> ResetFooterSegmentsForCompose() {
        _footerSegments = new System.Collections.Generic.List<FooterSegment>();
        return _footerSegments;
    }

    internal void ClearFooterSegmentsForCompose() {
        _footerSegments = null;
    }

    internal void SetFooterZonesForCompose(string? left, string? center, string? right) {
        ValidateZones(left, center, right, nameof(left));
        ClearFooterSegmentsForCompose();
        FooterFormat = string.Empty;
        ShowPageNumbers = true;
        _footerLeftFormat = left;
        _footerCenterFormat = center;
        _footerRightFormat = right;
    }

    internal void ClearFooterZonesForCompose() {
        _footerLeftFormat = null;
        _footerCenterFormat = null;
        _footerRightFormat = null;
    }

    internal void AddFooterImageForCompose(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        (_footerImages ??= new System.Collections.Generic.List<PdfHeaderFooterImage>()).Add(image.Clone());
    }

    internal void SetFirstPageFooterZonesForCompose(string? left, string? center, string? right) {
        ValidateZones(left, center, right, nameof(left));
        ClearFirstPageFooterSegmentsForCompose();
        DifferentFirstPageHeaderFooter = true;
        FirstPageFooterFormat = string.Empty;
        _firstPageFooterLeftFormat = left;
        _firstPageFooterCenterFormat = center;
        _firstPageFooterRightFormat = right;
    }

    internal void ClearFirstPageFooterZonesForCompose() {
        _firstPageFooterLeftFormat = null;
        _firstPageFooterCenterFormat = null;
        _firstPageFooterRightFormat = null;
    }

    internal void AddFirstPageFooterImageForCompose(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        DifferentFirstPageHeaderFooter = true;
        (_firstPageFooterImages ??= new System.Collections.Generic.List<PdfHeaderFooterImage>()).Add(image.Clone());
    }

    internal void SetEvenPageFooterZonesForCompose(string? left, string? center, string? right) {
        ValidateZones(left, center, right, nameof(left));
        ClearEvenPageFooterSegmentsForCompose();
        DifferentOddAndEvenPagesHeaderFooter = true;
        EvenPageFooterFormat = string.Empty;
        _evenPageFooterLeftFormat = left;
        _evenPageFooterCenterFormat = center;
        _evenPageFooterRightFormat = right;
    }

    internal void ClearEvenPageFooterZonesForCompose() {
        _evenPageFooterLeftFormat = null;
        _evenPageFooterCenterFormat = null;
        _evenPageFooterRightFormat = null;
    }

    internal void AddEvenPageFooterImageForCompose(PdfHeaderFooterImage image) {
        Guard.NotNull(image, nameof(image));
        DifferentOddAndEvenPagesHeaderFooter = true;
        (_evenPageFooterImages ??= new System.Collections.Generic.List<PdfHeaderFooterImage>()).Add(image.Clone());
    }

    internal void AddFooterShapeForCompose(PdfHeaderFooterShape shape) {
        Guard.NotNull(shape, nameof(shape));
        (_footerShapes ??= new System.Collections.Generic.List<PdfHeaderFooterShape>()).Add(shape.Clone());
    }

    internal void AddFirstPageFooterShapeForCompose(PdfHeaderFooterShape shape) {
        Guard.NotNull(shape, nameof(shape));
        DifferentFirstPageHeaderFooter = true;
        (_firstPageFooterShapes ??= new System.Collections.Generic.List<PdfHeaderFooterShape>()).Add(shape.Clone());
    }

    internal void AddEvenPageFooterShapeForCompose(PdfHeaderFooterShape shape) {
        Guard.NotNull(shape, nameof(shape));
        DifferentOddAndEvenPagesHeaderFooter = true;
        (_evenPageFooterShapes ??= new System.Collections.Generic.List<PdfHeaderFooterShape>()).Add(shape.Clone());
    }

    internal System.Collections.Generic.List<FooterSegment> ResetFirstPageFooterSegmentsForCompose() {
        _firstPageFooterSegments = new System.Collections.Generic.List<FooterSegment>();
        DifferentFirstPageHeaderFooter = true;
        return _firstPageFooterSegments;
    }

    internal void ClearFirstPageFooterSegmentsForCompose() {
        _firstPageFooterSegments = null;
    }

    internal System.Collections.Generic.List<FooterSegment> ResetEvenPageFooterSegmentsForCompose() {
        _evenPageFooterSegments = new System.Collections.Generic.List<FooterSegment>();
        DifferentOddAndEvenPagesHeaderFooter = true;
        return _evenPageFooterSegments;
    }

    internal void ClearEvenPageFooterSegmentsForCompose() {
        _evenPageFooterSegments = null;
    }

}

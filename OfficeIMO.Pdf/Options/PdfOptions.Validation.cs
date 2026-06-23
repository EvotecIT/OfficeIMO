namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    internal void Validate() {
        if (PageWidth <= 0 || double.IsNaN(PageWidth) || double.IsInfinity(PageWidth)) {
            throw new System.ArgumentException("PDF page width must be a positive finite value.");
        }

        if (PageHeight <= 0 || double.IsNaN(PageHeight) || double.IsInfinity(PageHeight)) {
            throw new System.ArgumentException("PDF page height must be a positive finite value.");
        }

        if (MarginLeft < 0 || double.IsNaN(MarginLeft) || double.IsInfinity(MarginLeft)) {
            throw new System.ArgumentException("PDF left margin must be a non-negative finite value.");
        }

        if (MarginRight < 0 || double.IsNaN(MarginRight) || double.IsInfinity(MarginRight)) {
            throw new System.ArgumentException("PDF right margin must be a non-negative finite value.");
        }

        if (MarginTop < 0 || double.IsNaN(MarginTop) || double.IsInfinity(MarginTop)) {
            throw new System.ArgumentException("PDF top margin must be a non-negative finite value.");
        }

        if (MarginBottom < 0 || double.IsNaN(MarginBottom) || double.IsInfinity(MarginBottom)) {
            throw new System.ArgumentException("PDF bottom margin must be a non-negative finite value.");
        }

        if (PageWidth - MarginLeft - MarginRight <= 0) {
            throw new System.ArgumentException("PDF margins must leave a positive content width.");
        }

        if (PageHeight - MarginTop - MarginBottom <= 0) {
            throw new System.ArgumentException("PDF margins must leave a positive content height.");
        }

        Guard.StandardFont(DefaultFont, nameof(DefaultFont), "PDF default font must be one of the supported standard PDF fonts.");
        Guard.StandardFont(HeaderFont, nameof(HeaderFont), "PDF header font must be one of the supported standard PDF fonts.");
        Guard.StandardFont(FooterFont, nameof(FooterFont), "PDF footer font must be one of the supported standard PDF fonts.");
        Guard.PageNumberStyle(PageNumberStyle, nameof(PageNumberStyle));
        Guard.ComplianceProfile(ComplianceProfile, nameof(ComplianceProfile));
        PdfPageLabelDictionaryBuilder.ValidatePrefix(PageLabelPrefix, nameof(PageLabelPrefix));
        if (_encryption != null && HasPdfABackedGroundwork()) {
            throw new System.ArgumentException("PDF Standard encryption cannot be combined with PDF/A, Factur-X, or ZUGFeRD groundwork.");
        }

        if (DefaultFontSize <= 0 || double.IsNaN(DefaultFontSize) || double.IsInfinity(DefaultFontSize)) {
            throw new System.ArgumentException("PDF default font size must be a positive finite value.");
        }

        if (HeaderFontSize <= 0 || double.IsNaN(HeaderFontSize) || double.IsInfinity(HeaderFontSize)) {
            throw new System.ArgumentException("PDF header font size must be a positive finite value.");
        }

        if (FooterFontSize <= 0 || double.IsNaN(FooterFontSize) || double.IsInfinity(FooterFontSize)) {
            throw new System.ArgumentException("PDF footer font size must be a positive finite value.");
        }

        if (HeaderOffsetY < 0 || double.IsNaN(HeaderOffsetY) || double.IsInfinity(HeaderOffsetY)) {
            throw new System.ArgumentException("PDF header offset must be a non-negative finite value.");
        }

        if (HasHeaderContent && HeaderOffsetY > MarginTop) {
            throw new System.ArgumentException("PDF header offset must not exceed the top margin when header content is enabled.");
        }

        if (HasHeaderContentForPage(1) && HeaderOffsetY > MarginTop) {
            throw new System.ArgumentException("PDF header offset must not exceed the top margin when header content is enabled.");
        }

        if (HasHeaderContentForPage(2) && HeaderOffsetY > MarginTop) {
            throw new System.ArgumentException("PDF header offset must not exceed the top margin when header content is enabled.");
        }

        if (FooterOffsetY < 0 || double.IsNaN(FooterOffsetY) || double.IsInfinity(FooterOffsetY)) {
            throw new System.ArgumentException("PDF footer offset must be a non-negative finite value.");
        }

        if (PageNumberStart < 1) {
            throw new System.ArgumentException("PDF page number start must be a positive value.");
        }

        if (HasFooterContent && FooterOffsetY > MarginBottom) {
            throw new System.ArgumentException("PDF footer offset must not exceed the bottom margin when footer content is enabled.");
        }

        if (HasFooterContentForPage(1) && FooterOffsetY > MarginBottom) {
            throw new System.ArgumentException("PDF footer offset must not exceed the bottom margin when footer content is enabled.");
        }

        if (HasFooterContentForPage(2) && FooterOffsetY > MarginBottom) {
            throw new System.ArgumentException("PDF footer offset must not exceed the bottom margin when footer content is enabled.");
        }

        if (HeaderFormat == null) {
            throw new System.ArgumentException("PDF header format cannot be null.");
        }

        if (FirstPageHeaderFormat == null) {
            throw new System.ArgumentException("PDF first-page header format cannot be null.");
        }

        if (EvenPageHeaderFormat == null) {
            throw new System.ArgumentException("PDF even-page header format cannot be null.");
        }

        if (FooterFormat == null) {
            throw new System.ArgumentException("PDF footer format cannot be null.");
        }

        if (FirstPageFooterFormat == null) {
            throw new System.ArgumentException("PDF first-page footer format cannot be null.");
        }

        if (EvenPageFooterFormat == null) {
            throw new System.ArgumentException("PDF even-page footer format cannot be null.");
        }

        ValidatePageTextSegments(_headerSegments, "header");
        ValidatePageTextSegments(_firstPageHeaderSegments, "header");
        ValidatePageTextSegments(_evenPageHeaderSegments, "header");
        ValidateFooterSegments(_footerSegments);
        ValidateFooterSegments(_firstPageFooterSegments);
        ValidateFooterSegments(_evenPageFooterSegments);
        ValidateZoneString(_headerLeftFormat, "header");
        ValidateZoneString(_headerCenterFormat, "header");
        ValidateZoneString(_headerRightFormat, "header");
        ValidateZoneString(_firstPageHeaderLeftFormat, "header");
        ValidateZoneString(_firstPageHeaderCenterFormat, "header");
        ValidateZoneString(_firstPageHeaderRightFormat, "header");
        ValidateZoneString(_evenPageHeaderLeftFormat, "header");
        ValidateZoneString(_evenPageHeaderCenterFormat, "header");
        ValidateZoneString(_evenPageHeaderRightFormat, "header");
        ValidateZoneString(_footerLeftFormat, "footer");
        ValidateZoneString(_footerCenterFormat, "footer");
        ValidateZoneString(_footerRightFormat, "footer");
        ValidateZoneString(_firstPageFooterLeftFormat, "footer");
        ValidateZoneString(_firstPageFooterCenterFormat, "footer");
        ValidateZoneString(_firstPageFooterRightFormat, "footer");
        ValidateZoneString(_evenPageFooterLeftFormat, "footer");
        ValidateZoneString(_evenPageFooterCenterFormat, "footer");
        ValidateZoneString(_evenPageFooterRightFormat, "footer");
        ValidateOptionalLanguage(Language, nameof(Language));
        ValidateOptionalCatalogUriBase(CatalogUriBase, nameof(CatalogUriBase));
        if (AcroFormDefaultTextAlignment.HasValue) {
            Guard.FormFieldTextAlignment(AcroFormDefaultTextAlignment.Value, nameof(AcroFormDefaultTextAlignment));
        }
    }

    private bool HasPdfABackedGroundwork() =>
        _pdfAIdentification != null ||
        _electronicInvoiceMetadata != null ||
        ComplianceProfile == PdfComplianceProfile.PdfA2B ||
        ComplianceProfile == PdfComplianceProfile.PdfA2U ||
        ComplianceProfile == PdfComplianceProfile.PdfA2A ||
        ComplianceProfile == PdfComplianceProfile.PdfA3B ||
        ComplianceProfile == PdfComplianceProfile.PdfA3U ||
        ComplianceProfile == PdfComplianceProfile.PdfA3A ||
        ComplianceProfile == PdfComplianceProfile.PdfA4 ||
        ComplianceProfile == PdfComplianceProfile.PdfA4E ||
        ComplianceProfile == PdfComplianceProfile.PdfA4F ||
        ComplianceProfile == PdfComplianceProfile.FacturX ||
        ComplianceProfile == PdfComplianceProfile.Zugferd;

    private static void ValidateZones(string? left, string? center, string? right, string paramName) {
        if (left == null && center == null && right == null) {
            throw new System.ArgumentException("At least one PDF header/footer zone must contain text.", paramName);
        }

        ValidateZoneString(left, "header/footer");
        ValidateZoneString(center, "header/footer");
        ValidateZoneString(right, "header/footer");
    }

    private static void ValidateZoneString(string? value, string scope) {
        if (value == null) {
            return;
        }

        if (value.Length == 0) {
            throw new System.ArgumentException("PDF " + scope + " zone text cannot be empty.");
        }
    }

    private static void ValidateOptionalLanguage(string? value, string paramName) {
        if (value == null) {
            return;
        }

        if (string.IsNullOrWhiteSpace(value)) {
            throw new System.ArgumentException("PDF document language cannot be empty or whitespace.", paramName);
        }

        for (int i = 0; i < value.Length; i++) {
            if (char.IsControl(value[i])) {
                throw new System.ArgumentException("PDF document language cannot contain control characters.", paramName);
            }
        }
    }

    private static void ValidateOptionalCatalogUriBase(string? value, string paramName) {
        if (value == null) {
            return;
        }

        if (string.IsNullOrWhiteSpace(value)) {
            throw new System.ArgumentException("PDF catalog URI base cannot be empty or whitespace.", paramName);
        }

        for (int i = 0; i < value.Length; i++) {
            if (char.IsControl(value[i])) {
                throw new System.ArgumentException("PDF catalog URI base cannot contain control characters.", paramName);
            }
        }

        if (!System.Uri.TryCreate(value, System.UriKind.Absolute, out _)) {
            throw new System.ArgumentException("PDF catalog URI base must be an absolute URI.", paramName);
        }
    }

    private static void ValidatePageTextSegments(System.Collections.Generic.List<FooterSegment>? segments, string scope) {
        if (segments != null) {
            for (int i = 0; i < segments.Count; i++) {
                var segment = segments[i];
                if (segment == null) {
                    throw new System.ArgumentException("PDF " + scope + " segments cannot contain null entries.");
                }

                if (segment.Kind == FooterSegmentKind.Text && segment.Text == null) {
                    throw new System.ArgumentException("PDF " + scope + " text segments cannot be null.");
                }

                if (segment.Kind != FooterSegmentKind.Text &&
                    segment.Kind != FooterSegmentKind.PageNumber &&
                    segment.Kind != FooterSegmentKind.TotalPages) {
                    throw new System.ArgumentException("PDF " + scope + " segments must use a supported segment kind.");
                }
            }
        }
    }

    private static void ValidateFooterSegments(System.Collections.Generic.List<FooterSegment>? segments) {
        if (segments != null) {
            for (int i = 0; i < segments.Count; i++) {
                var segment = segments[i];
                if (segment == null) {
                    throw new System.ArgumentException("PDF footer segments cannot contain null entries.");
                }

                if (segment.Kind == FooterSegmentKind.Text && segment.Text == null) {
                    throw new System.ArgumentException("PDF footer text segments cannot be null.");
                }

                if (segment.Kind != FooterSegmentKind.Text &&
                    segment.Kind != FooterSegmentKind.PageNumber &&
                    segment.Kind != FooterSegmentKind.TotalPages) {
                    throw new System.ArgumentException("PDF footer segments must use a supported segment kind.");
                }
            }
        }
    }
}

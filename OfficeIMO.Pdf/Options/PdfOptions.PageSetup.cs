namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    private bool _showFirstPageHeader = true;
    private bool _showEvenPageHeader = true;
    private bool _showFirstPageFooter = true;
    private bool _showEvenPageFooter = true;

    /// <summary>Page size, margins, orientation and running headers and footers of these options.</summary>
    public PdfPageSetup PageSetup => new PdfPageSetup(this);

    internal bool GetRunningTextShown(bool isHeader, PdfRunningContentVariant variant) => variant switch {
        PdfRunningContentVariant.FirstPage => isHeader ? _showFirstPageHeader : _showFirstPageFooter,
        PdfRunningContentVariant.EvenPages => isHeader ? _showEvenPageHeader : _showEvenPageFooter,
        _ => isHeader ? ShowHeader : ShowPageNumbers
    };

    internal void SetRunningTextShown(bool isHeader, PdfRunningContentVariant variant, bool value) {
        switch (variant) {
            case PdfRunningContentVariant.FirstPage:
                if (isHeader) _showFirstPageHeader = value; else _showFirstPageFooter = value;
                break;
            case PdfRunningContentVariant.EvenPages:
                if (isHeader) _showEvenPageHeader = value; else _showEvenPageFooter = value;
                break;
            default:
                if (isHeader) ShowHeader = value; else ShowPageNumbers = value;
                break;
        }
    }

    internal string GetRunningTextFormat(bool isHeader, PdfRunningContentVariant variant) => variant switch {
        PdfRunningContentVariant.FirstPage => isHeader ? FirstPageHeaderFormat : FirstPageFooterFormat,
        PdfRunningContentVariant.EvenPages => isHeader ? EvenPageHeaderFormat : EvenPageFooterFormat,
        _ => isHeader ? HeaderFormat : FooterFormat
    };

    internal void SetRunningTextFormat(bool isHeader, PdfRunningContentVariant variant, string value) {
        switch (variant) {
            case PdfRunningContentVariant.FirstPage:
                if (isHeader) FirstPageHeaderFormat = value; else FirstPageFooterFormat = value;
                break;
            case PdfRunningContentVariant.EvenPages:
                if (isHeader) EvenPageHeaderFormat = value; else EvenPageFooterFormat = value;
                break;
            default:
                if (isHeader) HeaderFormat = value; else FooterFormat = value;
                break;
        }
    }

    internal System.Collections.Generic.IReadOnlyList<PdfFooterSegment>? GetRunningTextSegments(bool isHeader, PdfRunningContentVariant variant) {
        System.Collections.Generic.List<PdfFooterSegment>? segments = variant switch {
            PdfRunningContentVariant.FirstPage => isHeader ? _firstPageHeaderSegments : _firstPageFooterSegments,
            PdfRunningContentVariant.EvenPages => isHeader ? _evenPageHeaderSegments : _evenPageFooterSegments,
            _ => isHeader ? _headerSegments : _footerSegments
        };
        return segments == null ? null : new System.Collections.Generic.List<PdfFooterSegment>(segments);
    }

    internal void SetRunningTextSegments(bool isHeader, PdfRunningContentVariant variant, System.Collections.Generic.IEnumerable<PdfFooterSegment>? value) {
        System.Collections.Generic.List<PdfFooterSegment>? copy = value == null ? null : new System.Collections.Generic.List<PdfFooterSegment>(value);
        switch (variant) {
            case PdfRunningContentVariant.FirstPage:
                if (isHeader) _firstPageHeaderSegments = copy; else _firstPageFooterSegments = copy;
                break;
            case PdfRunningContentVariant.EvenPages:
                if (isHeader) _evenPageHeaderSegments = copy; else _evenPageFooterSegments = copy;
                break;
            default:
                if (isHeader) _headerSegments = copy; else _footerSegments = copy;
                break;
        }
    }

    internal void SetRunningText(bool isHeader, PdfRunningContentVariant variant, string format) {
        switch (variant) {
            case PdfRunningContentVariant.FirstPage:
                if (isHeader) {
                    ClearFirstPageHeaderSegmentsForCompose();
                    ClearFirstPageHeaderZonesForCompose();
                    DifferentFirstPageHeaderFooter = true;
                    FirstPageHeaderFormat = format;
                } else {
                    ClearFirstPageFooterSegmentsForCompose();
                    ClearFirstPageFooterZonesForCompose();
                    DifferentFirstPageHeaderFooter = true;
                    FirstPageFooterFormat = format;
                }
                break;
            case PdfRunningContentVariant.EvenPages:
                if (isHeader) {
                    ClearEvenPageHeaderSegmentsForCompose();
                    ClearEvenPageHeaderZonesForCompose();
                    DifferentOddAndEvenPagesHeaderFooter = true;
                    EvenPageHeaderFormat = format;
                } else {
                    ClearEvenPageFooterSegmentsForCompose();
                    ClearEvenPageFooterZonesForCompose();
                    DifferentOddAndEvenPagesHeaderFooter = true;
                    EvenPageFooterFormat = format;
                }
                break;
            default:
                if (isHeader) {
                    ClearHeaderSegmentsForCompose();
                    ClearHeaderZonesForCompose();
                    ShowHeader = true;
                    HeaderFormat = format;
                } else {
                    ClearFooterSegmentsForCompose();
                    ClearFooterZonesForCompose();
                    ShowPageNumbers = true;
                    FooterFormat = format;
                }
                break;
        }
    }

    internal void SetRunningText(bool isHeader, PdfRunningContentVariant variant, System.Collections.Generic.IEnumerable<PdfFooterSegment> segments) {
        System.Collections.Generic.List<PdfFooterSegment> target;
        switch (variant) {
            case PdfRunningContentVariant.FirstPage:
                if (isHeader) {
                    target = ResetFirstPageHeaderSegmentsForCompose();
                    ClearFirstPageHeaderZonesForCompose();
                } else {
                    target = ResetFirstPageFooterSegmentsForCompose();
                    ClearFirstPageFooterZonesForCompose();
                }
                break;
            case PdfRunningContentVariant.EvenPages:
                if (isHeader) {
                    target = ResetEvenPageHeaderSegmentsForCompose();
                    ClearEvenPageHeaderZonesForCompose();
                } else {
                    target = ResetEvenPageFooterSegmentsForCompose();
                    ClearEvenPageFooterZonesForCompose();
                }
                break;
            default:
                if (isHeader) {
                    target = ResetHeaderSegmentsForCompose();
                    ClearHeaderZonesForCompose();
                } else {
                    target = ResetFooterSegmentsForCompose();
                    ClearFooterZonesForCompose();
                    ShowPageNumbers = true;
                }
                break;
        }
        foreach (PdfFooterSegment segment in segments) {
            Guard.NotNull(segment, nameof(segments));
            target.Add(segment);
        }
    }

    internal void SetRunningZones(bool isHeader, PdfRunningContentVariant variant, string? left, string? center, string? right) {
        switch (variant) {
            case PdfRunningContentVariant.FirstPage:
                if (isHeader) SetFirstPageHeaderZonesForCompose(left, center, right); else SetFirstPageFooterZonesForCompose(left, center, right);
                break;
            case PdfRunningContentVariant.EvenPages:
                if (isHeader) SetEvenPageHeaderZonesForCompose(left, center, right); else SetEvenPageFooterZonesForCompose(left, center, right);
                break;
            default:
                if (isHeader) SetHeaderZonesForCompose(left, center, right); else SetFooterZonesForCompose(left, center, right);
                break;
        }
    }

    internal void SetRunningZoneSegments(bool isHeader, PdfRunningContentVariant variant,
        System.Collections.Generic.List<PdfFooterSegment>? left,
        System.Collections.Generic.List<PdfFooterSegment>? center,
        System.Collections.Generic.List<PdfFooterSegment>? right) {
        switch (variant) {
            case PdfRunningContentVariant.FirstPage:
                if (isHeader) SetFirstPageHeaderZoneSegmentsForCompose(left, center, right); else SetFirstPageFooterZoneSegmentsForCompose(left, center, right);
                break;
            case PdfRunningContentVariant.EvenPages:
                if (isHeader) SetEvenPageHeaderZoneSegmentsForCompose(left, center, right); else SetEvenPageFooterZoneSegmentsForCompose(left, center, right);
                break;
            default:
                if (isHeader) SetHeaderZoneSegmentsForCompose(left, center, right); else SetFooterZoneSegmentsForCompose(left, center, right);
                break;
        }
    }

    internal void AddRunningImage(bool isHeader, PdfRunningContentVariant variant, PdfHeaderFooterImage image) {
        switch (variant) {
            case PdfRunningContentVariant.FirstPage:
                if (isHeader) AddFirstPageHeaderImageForCompose(image); else AddFirstPageFooterImageForCompose(image);
                break;
            case PdfRunningContentVariant.EvenPages:
                if (isHeader) AddEvenPageHeaderImageForCompose(image); else AddEvenPageFooterImageForCompose(image);
                break;
            default:
                if (isHeader) AddHeaderImageForCompose(image); else AddFooterImageForCompose(image);
                break;
        }
    }

    internal void AddRunningShape(bool isHeader, PdfRunningContentVariant variant, PdfHeaderFooterShape shape) {
        switch (variant) {
            case PdfRunningContentVariant.FirstPage:
                if (isHeader) AddFirstPageHeaderShapeForCompose(shape); else AddFirstPageFooterShapeForCompose(shape);
                break;
            case PdfRunningContentVariant.EvenPages:
                if (isHeader) AddEvenPageHeaderShapeForCompose(shape); else AddEvenPageFooterShapeForCompose(shape);
                break;
            default:
                if (isHeader) AddHeaderShapeForCompose(shape); else AddFooterShapeForCompose(shape);
                break;
        }
    }
}

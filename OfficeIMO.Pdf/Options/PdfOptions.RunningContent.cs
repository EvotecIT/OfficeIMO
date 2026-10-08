namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    // Geometry-only views live inside one serialization. Borrow its immutable assets and
    // generation font programs so story glyph usage reaches the page's font subset/CMap.
    // Public Clone remains an independent deep snapshot.
    internal PdfOptions CreateRunningContentFrame(double top, double bottom) {
        _embeddedFontPrograms ??= new Dictionary<PdfStandardFont, PdfTrueTypeFontProgram>();
        _embeddedOpenTypeCffFontPrograms ??= new Dictionary<PdfStandardFont, PdfOpenTypeCffFontProgram>();
        _namedFontPrograms ??= new Dictionary<PdfNamedFontFace, PdfTrueTypeFontProgram>();
        _namedOpenTypeCffFontPrograms ??= new Dictionary<PdfNamedFontFace, PdfOpenTypeCffFontProgram>();
        _namedFontProgramFailures ??= new HashSet<PdfNamedFontFace>();
        _embeddedFontProgramFailures ??= new HashSet<PdfStandardFont>();
        var frame = (PdfOptions)MemberwiseClone();
        frame.MarginTop = top;
        frame.MarginBottom = bottom;
        return frame;
    }

    internal bool HasAnyRunningContent => HeaderContent != null || FirstPageHeaderContent != null || EvenPageHeaderContent != null ||
        FooterContent != null || FirstPageFooterContent != null || EvenPageFooterContent != null;
    internal PdfRunningContent? HeaderContent { get; private set; }
    internal PdfRunningContent? FirstPageHeaderContent { get; private set; }
    internal PdfRunningContent? EvenPageHeaderContent { get; private set; }
    internal PdfRunningContent? FooterContent { get; private set; }
    internal PdfRunningContent? FirstPageFooterContent { get; private set; }
    internal PdfRunningContent? EvenPageFooterContent { get; private set; }

    internal PdfRunningContent? GetRunningContentForPage(int variantPageNumber, bool isHeader) {
        if (variantPageNumber == 1 && DifferentFirstPageHeaderFooter)
            return isHeader ? FirstPageHeaderContent : FirstPageFooterContent;
        if (IsEvenPageVariant(variantPageNumber))
            return isHeader ? EvenPageHeaderContent : EvenPageFooterContent;
        return isHeader ? HeaderContent : FooterContent;
    }

    internal void SetRunningContent(PdfRunningContent? content, bool isHeader, PdfRunningContentVariant variant) {
        if (variant == PdfRunningContentVariant.FirstPage) {
            if (content != null) DifferentFirstPageHeaderFooter = true;
            if (isHeader) FirstPageHeaderContent = content;
            else FirstPageFooterContent = content;
        } else if (variant == PdfRunningContentVariant.EvenPages) {
            if (content != null) DifferentOddAndEvenPagesHeaderFooter = true;
            if (isHeader) EvenPageHeaderContent = content;
            else EvenPageFooterContent = content;
        } else {
            if (isHeader) HeaderContent = content;
            else FooterContent = content;
        }
    }
}

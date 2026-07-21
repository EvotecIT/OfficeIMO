namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    /// <summary>
    /// Clones reusable rendering configuration while removing document-level content that must not leak into a canvas stamp.
    /// </summary>
    internal PdfOptions CloneForCanvasStampRendering() {
        PdfOptions clone = Clone();

        clone.ShowHeader = false;
        clone.ShowPageNumbers = false;
        clone.DifferentFirstPageHeaderFooter = false;
        clone.DifferentOddAndEvenPagesHeaderFooter = false;
        clone.HeaderFormat = string.Empty;
        clone.FirstPageHeaderFormat = string.Empty;
        clone.EvenPageHeaderFormat = string.Empty;
        clone.FooterFormat = string.Empty;
        clone.FirstPageFooterFormat = string.Empty;
        clone.EvenPageFooterFormat = string.Empty;
        clone._headerSegments = null;
        clone._firstPageHeaderSegments = null;
        clone._evenPageHeaderSegments = null;
        clone._footerSegments = null;
        clone._firstPageFooterSegments = null;
        clone._evenPageFooterSegments = null;
        clone._headerLeftFormat = null;
        clone._headerCenterFormat = null;
        clone._headerRightFormat = null;
        clone._firstPageHeaderLeftFormat = null;
        clone._firstPageHeaderCenterFormat = null;
        clone._firstPageHeaderRightFormat = null;
        clone._evenPageHeaderLeftFormat = null;
        clone._evenPageHeaderCenterFormat = null;
        clone._evenPageHeaderRightFormat = null;
        clone._footerLeftFormat = null;
        clone._footerCenterFormat = null;
        clone._footerRightFormat = null;
        clone._firstPageFooterLeftFormat = null;
        clone._firstPageFooterCenterFormat = null;
        clone._firstPageFooterRightFormat = null;
        clone._evenPageFooterLeftFormat = null;
        clone._evenPageFooterCenterFormat = null;
        clone._evenPageFooterRightFormat = null;
        clone._headerImages = null;
        clone._firstPageHeaderImages = null;
        clone._evenPageHeaderImages = null;
        clone._footerImages = null;
        clone._firstPageFooterImages = null;
        clone._evenPageFooterImages = null;
        clone._headerShapes = null;
        clone._firstPageHeaderShapes = null;
        clone._evenPageHeaderShapes = null;
        clone._footerShapes = null;
        clone._firstPageFooterShapes = null;
        clone._evenPageFooterShapes = null;

        clone.BackgroundColor = null;
        clone._textWatermark = null;
        clone._firstPageTextWatermark = null;
        clone._evenPageTextWatermark = null;
        clone._suppressFirstPageTextWatermark = false;
        clone._suppressEvenPageTextWatermark = false;
        clone._imageWatermark = null;
        clone._firstPageImageWatermark = null;
        clone._evenPageImageWatermark = null;
        clone._suppressFirstPageImageWatermark = false;
        clone._suppressEvenPageImageWatermark = false;
        clone._pageBorder = null;
        clone._pageBackgroundImage = null;
        clone._pageBackgroundShapes = null;
        clone._encryption = null;

        return clone;
    }
}

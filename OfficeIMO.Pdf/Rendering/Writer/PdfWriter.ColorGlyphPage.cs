using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Resources of the page that content is written for. A PNG color glyph becomes a decorative page image
    // drawn at its token; its text stays in the invisible, extractable run. Translucent palette colors share
    // one ExtGState per alpha with the page's other opacity states.
    private sealed class PageColorGlyphs : IColorGlyphPage {
        private readonly Func<LayoutResult.Page?> _page;

        internal PageColorGlyphs(Func<LayoutResult.Page?> page) => _page = page;

        internal PageColorGlyphs(LayoutResult.Page page) : this(() => page) {
        }

        public string? PlaceImage(OfficeBitmapGlyphImage image, OfficeTransform unitSquare) {
            LayoutResult.Page? page = _page();
            if (page == null) return null;
            OfficeImageInfo info = OfficeImageReader.Identify(image.Png);
            if (info.Format != OfficeImageFormat.Png || info.Width <= 0 || info.Height <= 0) return null;
            // Bounds for consumers of the placement box, such as tagging: the transformed square's extent.
            double minX = double.MaxValue, minY = double.MaxValue, maxX = double.MinValue, maxY = double.MinValue;
            foreach ((double u, double v) in new[] { (0D, 0D), (1D, 0D), (0D, 1D), (1D, 1D) }) {
                OfficePoint corner = unitSquare.TransformPoint(new OfficePoint(u, v));
                minX = Math.Min(minX, corner.X); maxX = Math.Max(maxX, corner.X);
                minY = Math.Min(minY, corner.Y); maxY = Math.Max(maxY, corner.Y);
            }
            var pageImage = new PageImage {
                Data = image.Png, Info = info, X = minX, Y = minY, W = maxX - minX, H = maxY - minY,
                UnitSquareTransform = unitSquare, IsInlineDecoration = true, Interpolate = true
            };
            page.Images.Add(pageImage);
            pageImage.InlineDrawToken = AllocateInlineImageDrawToken(page);
            return pageImage.InlineDrawToken;
        }

        public string? AlphaState(double alpha) {
            LayoutResult.Page? page = _page();
            return page == null ? null : EnsureHeaderFooterGraphicsState(page, alpha, alpha);
        }
    }
}

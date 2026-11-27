using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private static string? GetNativeImageLink(WordImage image) {
            Uri? uri = image.Hyperlink;
            if (uri == null && image.DrawingNode().FirstDescendant(WordXml.A, "hlinkClick").Attribute(WordXml.R, "id") is { } relationshipId) {
                uri = image.Paragraph.ExternalHyperlinkTarget(relationshipId);
            }

            return uri?.IsAbsoluteUri == true && OfficeDrawingLinkPolicy.TryNormalize(uri.AbsoluteUri, out string link)
                ? link : null;
        }

        // Word lays an inline picture out with its effect extent, the margin Word records around the
        // drawing for effects and for rounding the picture to whole pixels. The bottom extent lifts
        // the picture above the baseline, so it makes the line taller like the picture itself.
        private static double GetNativeInlineImageBottomExtent(WordImage image) {
            long? bottom = image.DrawingNode().Child(WordXml.Wp, "inline").Child(WordXml.Wp, "effectExtent").Int64Attr("b");
            return bottom is > 0L and <= 12_700L * 1_584L ? bottom.Value / 12_700D : 0D;
        }

        private static bool TryCreateNativeCellInlineImage(WordImage image, out PdfCore.PdfTextRun? run, WordPdfExportOptions? options = null,
            System.Threading.CancellationToken cancellationToken = default, bool reportPreparationWarnings = true,
            bool clipToParagraphFragment = false, string? linkUri = null) {
            run = null;
            if (!TryPrepareNativeBodyImage(image, reportPreparationWarnings ? options : null, "table cell image", out byte[] prepared,
                    cancellationToken: options?.CancellationToken ?? cancellationToken, svgFallbackReport: options)) return false;
            double width = image.Width?.ToPoints() ?? 144D;
            double height = image.Height?.ToPoints() ?? 144D;
            run = PdfCore.PdfTextRun.Inline(new PdfCore.PdfInlineImage(prepared, global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height),
                CreateNativeImageStyle(image, options, "inline image"), baselineOffset: global::OfficeIMO.OfficeLength.Points(GetNativeInlineImageBottomExtent(image)),
                linkUri: linkUri) {
                ClipToParagraphFragment = clipToParagraphFragment
            });
            return true;
        }
    }
}

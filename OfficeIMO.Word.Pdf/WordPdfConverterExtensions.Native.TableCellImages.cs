using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static string? GetNativeTableCellImageLink(WordImage image) {
            Uri? uri = image.Hyperlink;
            if (uri == null && image._Image?.Descendants<A.HyperlinkOnClick>().FirstOrDefault()?.Id?.Value is { } relationshipId) {
                WordParagraph paragraph = image.Paragraph;
                DocumentFormat.OpenXml.Packaging.HyperlinkRelationship? relationship = WordHyperLink.FindHyperlinkRelationship(
                    WordHyperLink.ResolveRelationshipPart(paragraph._document, paragraph._paragraph), relationshipId);
                uri = relationship is { IsExternal: true } ? relationship.Uri : null;
            }

            return uri?.IsAbsoluteUri == true && OfficeDrawingLinkPolicy.TryNormalize(uri.AbsoluteUri, out string link)
                ? link : null;
        }

        private static bool TryCreateNativeCellInlineImage(WordImage image, out PdfCore.PdfTextRun? run, WordPdfExportOptions? options = null,
            System.Threading.CancellationToken cancellationToken = default, bool reportPreparationWarnings = true,
            bool clipToParagraphFragment = false, string? linkUri = null) {
            run = null;
            if (!TryPrepareNativeBodyImage(image, reportPreparationWarnings ? options : null, "table cell image", out byte[] prepared,
                    cancellationToken: options?.CancellationToken ?? cancellationToken)) return false;
            double width = image.Width?.ToPoints() ?? 144D;
            double height = image.Height?.ToPoints() ?? 144D;
            run = PdfCore.PdfTextRun.Inline(new PdfCore.PdfInlineImage(prepared, global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height),
                CreateNativeImageStyle(image, options, "inline image"), baselineOffset: global::OfficeIMO.OfficeLength.Points(0D),
                linkUri: linkUri) {
                ClipToParagraphFragment = clipToParagraphFragment
            });
            return true;
        }
    }
}

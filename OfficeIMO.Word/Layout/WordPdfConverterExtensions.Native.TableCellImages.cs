using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool TryCreateNativeCellInlineImage(WordImage image, out PdfCore.PdfTextRun? run, WordToPdfOptions? options = null,
            System.Threading.CancellationToken cancellationToken = default) {
            run = null;
            if (!TryPrepareNativeBodyImage(image, options, "table cell image", out byte[] prepared,
                    cancellationToken: cancellationToken)) return false;
            double width = image.Width.HasValue ? image.Width.Value * 72D / 96D : 144D;
            double height = image.Height.HasValue ? image.Height.Value * 72D / 96D : 144D;
            run = PdfCore.PdfTextRun.Inline(new PdfCore.PdfInlineImage(prepared, width, height,
                CreateNativeImageStyle(image, options, "inline image"), baselineOffset: 0D));
            return true;
        }
    }
}

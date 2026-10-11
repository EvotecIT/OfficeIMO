using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool TryCreateNativeCellInlineImage(WordImage image, out PdfCore.PdfTextRun? run, WordToPdfOptions? options = null) {
            run = null;
            if (!TryGetNativeBodyImageBytes(image, options, "table cell image", out byte[] bytes)) return false;
            if (!TryPrepareNativePdfImageBytes(bytes, out byte[] prepared, out _)) return false;
            double width = image.Width.HasValue ? image.Width.Value * 72D / 96D : 144D;
            double height = image.Height.HasValue ? image.Height.Value * 72D / 96D : 144D;
            string? alternativeText = string.IsNullOrWhiteSpace(image.Description) ? null : image.Description;
            run = PdfCore.PdfTextRun.Inline(new PdfCore.PdfInlineImage(prepared, width, height,
                alternativeText: alternativeText, fit: OfficeImageFit.Stretch));
            return true;
        }
    }
}

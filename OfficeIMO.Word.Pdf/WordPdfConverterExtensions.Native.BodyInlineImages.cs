using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static IReadOnlyDictionary<OpenXmlElement, PdfCore.PdfTextRun> CreateNativeBodyInlineImages(
            WordParagraph paragraph, IReadOnlyList<WordParagraph> runs, bool supported, WordToPdfOptions? options) {
            var images = new Dictionary<OpenXmlElement, PdfCore.PdfTextRun>();
            if (!supported) return images;
            int limit = options?.MaxImagesPerParagraph ?? 1_000;
            if (limit <= 0) throw new ArgumentOutOfRangeException(nameof(WordToPdfOptions.MaxImagesPerParagraph));
            int count = 0;
            foreach (WordParagraph run in runs) {
                if (IsNativeHiddenTextRun(run, paragraph)) continue;
                foreach (WordImage image in run.EnumerateImages()) {
                    options?.CancellationToken.ThrowIfCancellationRequested();
                    if (++count > limit) throw new InvalidDataException("Word paragraph image count exceeds the PDF export limit.");
                    if (image._Image?.Inline != null && !image._Image.Ancestors<W.SdtRun>().Any(IsNativePictureControl) &&
                        ReferenceEquals(image._Image.Ancestors<W.TextBoxContent>().FirstOrDefault(),
                            paragraph._paragraph?.Ancestors<W.TextBoxContent>().FirstOrDefault()) &&
                        TryCreateNativeBodyInlineImage(image, out PdfCore.PdfTextRun? inline))
                        images[image._Image] = inline!;
                }
            }
            return images;
        }
        private static bool TryCreateNativeBodyInlineImage(WordImage image, out PdfCore.PdfTextRun? inline) {
            try { return TryCreateNativeCellInlineImage(image, out inline); }
            catch (InvalidOperationException) {
                // The ordinary image path owns unavailable/linked-image diagnostics.
                inline = null;
                return false;
            }
        }
    }
}

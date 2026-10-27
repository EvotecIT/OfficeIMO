using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static IReadOnlyDictionary<OpenXmlElement, PdfCore.PdfTextRun> CreateNativeFlowInlineImages(
            WordParagraph paragraph, IReadOnlyList<WordParagraph> runs, bool supported, WordPdfExportOptions? options) {
            var images = new Dictionary<OpenXmlElement, PdfCore.PdfTextRun>();
            if (!supported) return images;
            // Main and note stories retain their imported overflow policy.
            // DrawingML running pictures use a picture-only exact-line clip.
            // Ordinary inline VML retains full paint in the same line box.
            // Cells and text boxes keep their existing policies.
            bool ordinaryStoryParagraph = paragraph._paragraph != null &&
                !paragraph._paragraph.Ancestors<W.TableCell>().Any() &&
                !paragraph._paragraph.Ancestors<W.TextBoxContent>().Any();
            bool allowFrameOverflow = ordinaryStoryParagraph && paragraph._paragraph!.Ancestors().Any(story =>
                story is W.Body or W.FootnoteEndnoteType);
            bool clipToParagraphFragment = ordinaryStoryParagraph && paragraph._paragraph!.Ancestors().Any(story =>
                story is W.Header or W.Footer);
            if (!allowFrameOverflow && !runs.Any(run => IsNativeRenderableTextRun(run, paragraph))) return images;
            int limit = options?.MaxImagesPerParagraph ?? 1_000;
            if (limit <= 0) throw new ArgumentOutOfRangeException(nameof(WordPdfExportOptions.MaxImagesPerParagraph));
            int count = 0;
            foreach (WordParagraph run in runs) {
                if (IsNativeHiddenTextRun(run, paragraph)) continue;
                foreach (WordImage image in run.EnumerateImages()) {
                    options?.CancellationToken.ThrowIfCancellationRequested();
                    if (++count > limit) throw new InvalidDataException("Word paragraph image count exceeds the PDF export limit.");
                    OpenXmlElement occurrence = image.OccurrenceElement;
                    bool inlineVml = clipToParagraphFragment && IsNativeInlineVmlPicture(image._vmlShape);
                    if ((image._Image?.Inline != null || inlineVml) && !occurrence.Ancestors<W.SdtRun>().Any(IsNativePictureControl) &&
                        ReferenceEquals(occurrence.Ancestors<W.TextBoxContent>().FirstOrDefault(),
                            paragraph._paragraph?.Ancestors<W.TextBoxContent>().FirstOrDefault()) &&
                        TryCreateNativeFlowInlineImage(image, out PdfCore.PdfTextRun? inline, options,
                            options?.CancellationToken ?? default, allowFrameOverflow, clipToParagraphFragment && !inlineVml))
                        images[occurrence] = inline!;
                }
            }
            return images;
        }
        private static bool TryCreateNativeFlowInlineImage(WordImage image, out PdfCore.PdfTextRun? inline, WordPdfExportOptions? options,
            System.Threading.CancellationToken cancellationToken, bool allowFrameOverflow, bool clipToParagraphFragment) {
            try {
                if (!allowFrameOverflow)
                    return TryCreateNativeCellInlineImage(image, out inline, options, cancellationToken: cancellationToken,
                        reportPreparationWarnings: false, clipToParagraphFragment: clipToParagraphFragment);
                inline = null;
                // The ordinary paragraph renderer owns unavailable-image diagnostics.
                // A successful probe still maps crop diagnostics through options.
                if (!TryPrepareNativeBodyImage(image, null, "flow inline image", out byte[] prepared,
                        cancellationToken: cancellationToken)) return false;
                double width = image.Width?.ToPoints() ?? 144D;
                double height = image.Height?.ToPoints() ?? 144D;
                PdfCore.PdfImageStyle style = CreateNativeImageStyle(image, options, "inline image");
                style.ScaleDownToFit = false;
                inline = PdfCore.PdfTextRun.Inline(new PdfCore.PdfInlineImage(prepared, width, height, style, 0D) {
                    AllowFrameOverflow = true
                });
                return true;
            } catch (InvalidOperationException) {
                // The ordinary image path owns unavailable/linked-image diagnostics.
                inline = null;
                return false;
            }
        }
    }
}

using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private static IReadOnlyDictionary<OfficeXmlNode, PdfCore.PdfTextRun> CreateNativeFlowInlineImages(
            WordParagraph paragraph, IReadOnlyList<WordParagraph> runs, bool supported, WordPdfExportOptions? options) {
            var images = new Dictionary<OfficeXmlNode, PdfCore.PdfTextRun>();
            if (!supported) return images;
            // Main and note stories retain their imported overflow policy.
            // DrawingML running pictures use a picture-only exact-line clip.
            // Ordinary inline VML retains full paint in the same line box.
            // Cells and text boxes keep their existing policies.
            OfficeXmlNode paragraphNode = paragraph.Node();
            bool ordinaryStoryParagraph = !paragraphNode.IsNone &&
                paragraphNode.Ancestor(WordXml.W, "tc").IsNone &&
                paragraphNode.Ancestor(WordXml.W, "txbxContent").IsNone;
            bool allowFrameOverflow = ordinaryStoryParagraph && paragraphNode.Ancestors().Any(story =>
                story.IsW("body") || story.IsW("footnote") || story.IsW("endnote"));
            bool clipToParagraphFragment = ordinaryStoryParagraph && paragraphNode.Ancestors().Any(story =>
                story.IsW("hdr") || story.IsW("ftr"));
            if (!allowFrameOverflow && !runs.Any(run => IsNativeRenderableTextRun(run, paragraph))) return images;
            int limit = options?.MaxImagesPerParagraph ?? 1_000;
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(limit, nameof(WordPdfExportOptions.MaxImagesPerParagraph));
            int count = 0;
            foreach (WordParagraph run in runs) {
                if (IsNativeHiddenTextRun(run, paragraph)) continue;
                foreach (WordImage image in run.EnumerateImages()) {
                    options?.CancellationToken.ThrowIfCancellationRequested();
                    if (++count > limit) throw new InvalidDataException("Word paragraph image count exceeds the PDF export limit.");
                    OfficeXmlNode occurrence = image.OccurrenceNode();
                    bool inlineVml = clipToParagraphFragment &&
                        IsNativeInlineVmlPicture(image.VmlShapeNode());
                    if ((!image.DrawingNode().Child(WordXml.Wp, "inline").IsNone || inlineVml) &&
                        !occurrence.Ancestors().Any(ancestor => ancestor.IsRunContentControl() && IsNativePictureControl(ancestor)) &&
                        occurrence.Ancestor(WordXml.W, "txbxContent").Equals(paragraphNode.Ancestor(WordXml.W, "txbxContent")) &&
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
                        reportPreparationWarnings: false, clipToParagraphFragment: clipToParagraphFragment,
                        linkUri: GetNativeImageLink(image));
                inline = null;
                // The ordinary paragraph renderer owns unavailable-image diagnostics.
                // A successful probe still maps crop diagnostics through options.
                if (!TryPrepareNativeBodyImage(image, null, "flow inline image", out byte[] prepared,
                        cancellationToken: cancellationToken, svgFallbackReport: options)) return false;
                double width = image.Width?.ToPoints() ?? 144D;
                double height = image.Height?.ToPoints() ?? 144D;
                PdfCore.PdfImageStyle style = CreateNativeImageStyle(image, options, "inline image");
                style.ScaleDownToFit = false;
                inline = PdfCore.PdfTextRun.Inline(new PdfCore.PdfInlineImage(prepared, global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height), style, global::OfficeIMO.OfficeLength.Points(GetNativeInlineImageBottomExtent(image)),
                    linkUri: GetNativeImageLink(image)) {
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

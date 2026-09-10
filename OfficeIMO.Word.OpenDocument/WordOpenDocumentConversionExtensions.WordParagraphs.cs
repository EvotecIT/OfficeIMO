using OfficeIMO.OpenDocument;
using OfficeIMO.Word;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentConversionExtensions {
    private static void CopyParagraph(WordParagraphSnapshot source, OdtParagraph target,
        WordOpenDocumentConversionOptions options, OdfImageValidationBudget imageValidationBudget,
        ref int hyperlinks, ref int images, ref int unsupportedImages,
        ref int bookmarks, ref int unsupportedFootnotes) {
        bool wrote = false;
        OdtParagraph first = target;
        target.PageBreakBefore = source.PageBreakBefore;
        ApplyWordParagraphFormatting(source, target);
        foreach (WordRunSnapshot run in source.Runs) {
            int start = 0;
            int imageIndex = 0;
            if (run.NonTextBreaks != null) {
                foreach (var boundary in run.NonTextBreaks.OrderBy(item => item.Key)) {
                    AppendRunSegment(run, ref start, boundary.Key, ref imageIndex, target, options,
                        imageValidationBudget, ref hyperlinks, ref images, ref unsupportedImages, ref wrote);
                    if (boundary.Value == WordBreakType.Page) {
                        if (target.InlineNodes.Count > 0 || target.PageBreakBefore) {
                            target = target.InsertParagraphAfter();
                            ApplyWordParagraphFormatting(source, target);
                        }
                        target.PageBreakBefore = true;
                    } else AppendText(run, "\n", target, ref hyperlinks, ref wrote);
                    start = boundary.Key + 1;
                }
            }
            AppendRunSegment(run, ref start, run.Text.Length, ref imageIndex, target, options,
                imageValidationBudget, ref hyperlinks, ref images, ref unsupportedImages, ref wrote);
            if (run.Footnote != null) unsupportedFootnotes++;
        }
        if (!wrote && source.Text.Length > 0 && source.Runs.All(run => run.NonTextBreaks == null)) target.Text = source.Text;
        if (!string.IsNullOrWhiteSpace(source.BookmarkName)) { first.AddBookmark(source.BookmarkName!); bookmarks++; }
    }

    private static void AppendRunSegment(WordRunSnapshot run, ref int start, int end, ref int imageIndex,
        OdtParagraph target, WordOpenDocumentConversionOptions options, OdfImageValidationBudget imageValidationBudget,
        ref int hyperlinks, ref int images, ref int unsupportedImages, ref bool wrote) {
        while (imageIndex < run.PositionedImages.Count && run.PositionedImages[imageIndex].Offset <= end) {
            WordPositionedImageSnapshot positioned = run.PositionedImages[imageIndex++];
            AppendText(run, run.Text.Substring(start, positioned.Offset - start), target, ref hyperlinks, ref wrote);
            start = positioned.Offset;
            WordInlineImageSnapshot image = positioned.Image;
            if (options.IncludeImages && image.Bytes is { Length: > 0 } bytes) {
                try {
                    string fileName = image.FileName ?? "image.png";
                    if (!OdfImagePayloadValidator.TryResolvePreservedFileName(
                        bytes,
                        fileName,
                        out string storedFileName,
                        imageValidationBudget)) {
                        throw new NotSupportedException("The Word image payload is incomplete or unsupported.");
                    }
                    target.AddImage(bytes, storedFileName,
                        OdfLength.Points(image.Width ?? 72D), OdfLength.Points(image.Height ?? 72D),
                        image.IsInline ? OdtImageAnchor.Inline : OdtImageAnchor.Paragraph);
                    images++;
                    wrote = true;
                } catch (NotSupportedException) {
                    unsupportedImages++;
                }
            }
        }
        AppendText(run, run.Text.Substring(start, end - start), target, ref hyperlinks, ref wrote);
        start = end;
    }

    private static void AppendText(WordRunSnapshot run, string text, OdtParagraph target, ref int hyperlinks, ref bool wrote) {
        if (text.Length == 0) return;
        if (run.IsHyperlink && (!string.IsNullOrWhiteSpace(run.HyperlinkUri) || !string.IsNullOrWhiteSpace(run.HyperlinkAnchor))) {
            OdtHyperlink link = target.AddHyperlink(text, run.HyperlinkUri ?? "#" + run.HyperlinkAnchor);
            ApplyWordRunFormatting(run, link);
            hyperlinks++;
        } else ApplyWordRunFormatting(run, target.AddSpan(text));
        wrote = true;
    }
}

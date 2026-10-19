using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Matching adjacent paragraphs share the outer border and one optional separator.</summary>
    private static void JoinNativeParagraphBorderGroup(WordParagraph paragraph, NativeParagraphBorders borders,
        PdfCore.PdfPanelStyle style) {
        if (!HasNativeParagraphBorder(borders)) return;
        bool previous = HasNativeMatchingBorderNeighbor(paragraph, borders, next: false);
        bool next = HasNativeMatchingBorderNeighbor(paragraph, borders, next: true);
        if (!previous && !next) return;
        // Each explicit side is already mapped. A uniform fallback would restore
        // a top or bottom edge that the shared frame deliberately suppresses.
        style.BorderColor = null;
        style.BorderWidth = 0D;
        double separatorWidth = HasNativeBorder(borders.Between.Style) ? (borders.Between.Size ?? 4U) / 8D : 0D;
        if (previous) {
            style.TopBorder = null;
            W.Paragraph previousSource = (W.Paragraph)paragraph._paragraph!.PreviousSibling()!;
            NativeParagraphBorderSide previousSeparator = GetNativeEffectiveParagraphBorders(new WordParagraph(paragraph._document, previousSource)).Between;
            double previousWidth = HasNativeBorder(previousSeparator.Style) ? (previousSeparator.Size ?? 4U) / 8D : 0D;
            double previousSpace = HasNativeBorder(previousSeparator.Style)
                ? Math.Min(previousSeparator.Space ?? 0D, MaxNativeParagraphBorderSpacingPoints) : 0D;
            style.PaddingTopOverride = previousSpace + previousWidth / 2D;
            style.FrameStartPaddingTopOverride = previousSpace + previousWidth;
            style.FrameStartTopBorder = CreateNativePanelBorder(previousSeparator);
            if (style.FrameStartTopBorder is { } frameStartBorder) frameStartBorder.Offset = -previousWidth / 2D;
            // The group's final outer edge closes the joined frame. Unlike an
            // intermediate separator, its full closing decoration must fit.
            if (!next) style.AllowTrailingPaddingOverflow = false;
        }
        if (next) {
            W.Paragraph nextSource = (W.Paragraph)paragraph._paragraph!.NextSibling()!;
            NativeParagraphBorderSide nextSeparator = GetNativeEffectiveParagraphBorders(new WordParagraph(paragraph._document, nextSource)).Between;
            double nextSpace = HasNativeBorder(nextSeparator.Style)
                ? Math.Min(nextSeparator.Space ?? 0D, MaxNativeParagraphBorderSpacingPoints) : 0D;
            style.BorderContinuesWithNext = true;
            style.BottomBorder = CreateNativePanelBorder(borders.Between);
            if (style.BottomBorder is { } separator) separator.Offset = 0D;
            // The preceding paragraph supplies the stroke; the following
            // paragraph supplies the space above that stroke. The preceding
            // space is reserved below it by the following paragraph's panel.
            style.PaddingBottomOverride = nextSpace + separatorWidth / 2D;
            // Modern layout may overflow the trailing stroke, but reserves its
            // text-to-border space. Legacy layout also reserves half the stroke.
            style.FragmentBottomInset = UsesModernNativeWordLayout(paragraph._document)
                ? nextSpace : nextSpace + separatorWidth / 2D;
        }
    }

    private static bool HasNativeMatchingBorderNeighbor(WordParagraph paragraph, NativeParagraphBorders borders, bool next) {
        W.Paragraph? source = paragraph._paragraph;
        if (source == null || !HasNativeParagraphBorder(borders) || !CanJoinNativeParagraphBorderGroup(paragraph)) return false;
        if ((next ? source.NextSibling() : source.PreviousSibling()) is not W.Paragraph neighborSource) return false;
        var neighbor = new WordParagraph(paragraph._document, neighborSource);
        WordParagraph first = next ? paragraph : neighbor;
        WordParagraph second = next ? neighbor : paragraph;
        if (!CanJoinNativeParagraphBorderGroup(neighbor) || HasNativePageBreakBefore(second) ||
            first._paragraph?.ParagraphProperties?.SectionProperties != null ||
            HasNativeParagraphBorderFlowBreak(first) || HasNativeParagraphBorderFlowBreak(second)) return false;
        NativeParagraphBorders other = GetNativeEffectiveParagraphBorders(neighbor);
        if (!HasNativeParagraphBorder(other)) return false;
        if (!NativeParagraphBordersMatch(borders.Top, other.Top) || !NativeParagraphBordersMatch(borders.Right, other.Right) ||
            !NativeParagraphBordersMatch(borders.Bottom, other.Bottom) || !NativeParagraphBordersMatch(borders.Left, other.Left)) return false;
        NativeParagraphStyleDefaults currentDefaults = GetNativeParagraphStyleDefaults(paragraph);
        NativeParagraphStyleDefaults neighborDefaults = GetNativeParagraphStyleDefaults(neighbor);
        return ResolveNativeParagraphBorderGroupLeft(paragraph, currentDefaults) ==
                ResolveNativeParagraphBorderGroupLeft(neighbor, neighborDefaults) &&
            (paragraph.IndentationAfterPoints ?? currentDefaults.RightIndent ?? 0D) ==
                (neighbor.IndentationAfterPoints ?? neighborDefaults.RightIndent ?? 0D);
    }

    private static bool HasNativeParagraphBorderFlowBreak(WordParagraph paragraph) =>
        paragraph._paragraph?.Descendants<W.Break>().Any(b => b.Type?.Value is { } type &&
            (type == W.BreakValues.Page || type == W.BreakValues.Column)) == true;

    private static bool HasNativeBorderGroupText(W.Paragraph paragraph) =>
        paragraph.Descendants<W.Text>().Any(text => !string.IsNullOrEmpty(text.Text));

    private static bool CanJoinNativeParagraphBorderGroup(WordParagraph paragraph) =>
        paragraph._paragraph is { } source && HasNativeBorderGroupText(source) && GetHeadingLevel(paragraph) == 0 &&
        // Mixed text boxes emit separate flow segments. Numbering can use a list
        // block or change the frame indents after paragraph style resolution.
        // Neither route currently owns a compatible shared paragraph frame.
        !source.Descendants<W.TextBoxContent>().Any() && WordDocumentTraversal.GetListInfo(paragraph) == null;

    private static double ResolveNativeParagraphBorderGroupLeft(WordParagraph paragraph, NativeParagraphStyleDefaults defaults) {
        double firstLine = paragraph.IndentationHangingPoints is { } hanging ? -hanging :
            paragraph.IndentationFirstLinePoints ?? defaults.FirstLineIndent ?? 0D;
        return (paragraph.IndentationBeforePoints ?? defaults.LeftIndent ?? 0D) + Math.Min(firstLine, 0D);
    }

    private static bool NativeParagraphBordersMatch(NativeParagraphBorderSide first, NativeParagraphBorderSide second) {
        if (!HasNativeBorder(first.Style) || !HasNativeBorder(second.Style))
            return !HasNativeBorder(first.Style) && !HasNativeBorder(second.Style);
        return first.Style == second.Style && first.Shadow == second.Shadow && first.Frame == second.Frame &&
            (first.Size ?? 4U) == (second.Size ?? 4U) &&
            Math.Min(first.Space ?? 0U, 31U) == Math.Min(second.Space ?? 0U, 31U) &&
            string.Equals(NormalizeNativeBorderColor(first.ColorHex),
                NormalizeNativeBorderColor(second.ColorHex), StringComparison.OrdinalIgnoreCase);
    }
}

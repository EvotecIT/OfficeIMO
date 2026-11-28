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
        style.BorderWidth = global::OfficeIMO.OfficeLength.Points(0D);
        double separatorWidth = HasNativeBorder(borders.Between.Style) ? (borders.Between.Size ?? 4U) / 8D : 0D;
        if (previous) {
            style.TopBorder = null;
            OfficeXmlNode previousSource = paragraph.Node().PreviousSibling;
            NativeParagraphBorderSide previousSeparator = GetNativeEffectiveParagraphBorders(paragraph._document.ParagraphView(previousSource)).Between;
            double previousWidth = HasNativeBorder(previousSeparator.Style) ? (previousSeparator.Size ?? 4U) / 8D : 0D;
            double previousSpace = HasNativeBorder(previousSeparator.Style)
                ? Math.Min(previousSeparator.Space ?? 0D, MaxNativeParagraphBorderSpacingPoints) : 0D;
            style.PaddingTopOverride = previousSpace + previousWidth / 2D;
            style.FrameStartPaddingTopOverride = previousSpace + previousWidth;
            style.FrameStartTopBorder = CreateNativePanelBorder(previousSeparator);
            if (style.FrameStartTopBorder is { } frameStartBorder) frameStartBorder.Offset = global::OfficeIMO.OfficeLength.Points(-previousWidth / 2D);
            // The group's final outer edge closes the joined frame. Unlike an
            // intermediate separator, its full closing decoration must fit.
            if (!next) style.AllowTrailingPaddingOverflow = false;
        }
        if (next) {
            OfficeXmlNode nextSource = paragraph.Node().NextSibling;
            NativeParagraphBorderSide nextSeparator = GetNativeEffectiveParagraphBorders(paragraph._document.ParagraphView(nextSource)).Between;
            double nextSpace = HasNativeBorder(nextSeparator.Style)
                ? Math.Min(nextSeparator.Space ?? 0D, MaxNativeParagraphBorderSpacingPoints) : 0D;
            style.BorderContinuesWithNext = true;
            style.BottomBorder = CreateNativePanelBorder(borders.Between);
            if (style.BottomBorder is { } separator) separator.Offset = global::OfficeIMO.OfficeLength.Points(0D);
            // The preceding paragraph supplies the stroke; the following
            // paragraph supplies the space above that stroke. The preceding
            // space is reserved below it by the following paragraph's panel.
            style.PaddingBottomOverride = nextSpace + separatorWidth / 2D;
            // Modern layout may overflow the trailing stroke, but reserves its
            // text-to-border space. Legacy layout also reserves half the stroke.
            style.FragmentBottomInset = global::OfficeIMO.OfficeLength.Points(UsesModernNativeWordLayout(paragraph._document)
                ? nextSpace : nextSpace + separatorWidth / 2D);
        }
    }

    private static bool HasNativeMatchingBorderNeighbor(WordParagraph paragraph, NativeParagraphBorders borders, bool next) {
        OfficeXmlNode source = paragraph.Node();
        if (source.IsNone || !HasNativeParagraphBorder(borders) || !CanJoinNativeParagraphBorderGroup(paragraph)) return false;
        OfficeXmlNode neighborSource = next ? source.NextSibling : source.PreviousSibling;
        if (!neighborSource.IsW("p")) return false;
        WordParagraph neighbor = paragraph._document.ParagraphView(neighborSource);
        WordParagraph first = next ? paragraph : neighbor;
        WordParagraph second = next ? neighbor : paragraph;
        if (!CanJoinNativeParagraphBorderGroup(neighbor) || HasNativePageBreakBefore(second) ||
            first.Node().WChild("pPr").WChild("sectPr").Exists ||
            HasNativeParagraphBorderFlowBreak(first) || HasNativeParagraphBorderFlowBreak(second)) return false;
        NativeParagraphBorders other = GetNativeEffectiveParagraphBorders(neighbor);
        if (!HasNativeParagraphBorder(other)) return false;
        if (!NativeParagraphBordersMatch(borders.Top, other.Top) || !NativeParagraphBordersMatch(borders.Right, other.Right) ||
            !NativeParagraphBordersMatch(borders.Bottom, other.Bottom) || !NativeParagraphBordersMatch(borders.Left, other.Left)) return false;
        NativeParagraphStyleDefaults currentDefaults = GetNativeParagraphStyleDefaults(paragraph);
        NativeParagraphStyleDefaults neighborDefaults = GetNativeParagraphStyleDefaults(neighbor);
        return ResolveNativeParagraphBorderGroupLeft(paragraph, currentDefaults) ==
                ResolveNativeParagraphBorderGroupLeft(neighbor, neighborDefaults) &&
            (paragraph.IndentationAfter?.ToPoints() ?? currentDefaults.RightIndent ?? 0D) ==
                (neighbor.IndentationAfter?.ToPoints() ?? neighborDefaults.RightIndent ?? 0D);
    }

    private static bool HasNativeParagraphBorderFlowBreak(WordParagraph paragraph) =>
        paragraph.Node().DescendantElements(WordXml.W, "br").Any(b => b.WAttr("type") is "page" or "column");

    private static bool HasNativeBorderGroupText(OfficeXmlNode paragraph) =>
        paragraph.DescendantElements(WordXml.W, "t").Any(text => !string.IsNullOrEmpty(text.InnerText));

    private static bool CanJoinNativeParagraphBorderGroup(WordParagraph paragraph) =>
        paragraph.Node() is { Exists: true } source && HasNativeBorderGroupText(source) && GetHeadingLevel(paragraph) == 0 &&
        // Mixed text boxes emit separate flow segments. Numbering can use a list
        // block or change the frame indents after paragraph style resolution.
        // Neither route currently owns a compatible shared paragraph frame.
        !source.HasDescendant(WordXml.W, "txbxContent") && WordDocumentTraversal.GetListInfo(paragraph) == null;

    private static double ResolveNativeParagraphBorderGroupLeft(WordParagraph paragraph, NativeParagraphStyleDefaults defaults) {
        double firstLine = paragraph.IndentationHanging?.ToPoints() is { } hanging ? -hanging :
            paragraph.IndentationFirstLine?.ToPoints() ?? defaults.FirstLineIndent ?? 0D;
        return (paragraph.IndentationBefore?.ToPoints() ?? defaults.LeftIndent ?? 0D) + Math.Min(firstLine, 0D);
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

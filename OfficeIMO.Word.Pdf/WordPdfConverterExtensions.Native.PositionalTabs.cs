using DocumentFormat.OpenXml;
using PdfCore = OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static PdfCore.PdfTextRun CreateNativePositionalTab(W.PositionalTab tab, WordParagraph paragraph,
        NativeDocumentDefaults defaults, NativeTableStyleDefaults? tableDefaults = null) {
        string? alignment = tab.Alignment?.InnerText;
        double fraction = alignment == "center" ? 0.5D : alignment == "right" ? 1D : 0D;
        PdfCore.PdfTabAlignment pdfAlignment = alignment == "center" ? PdfCore.PdfTabAlignment.Center :
            alignment == "right" ? PdfCore.PdfTabAlignment.Right : PdfCore.PdfTabAlignment.Left;
        PdfCore.PdfTabLeaderStyle leader = tab.Leader?.InnerText switch {
            "dot" => PdfCore.PdfTabLeaderStyle.Dots,
            "hyphen" => PdfCore.PdfTabLeaderStyle.Hyphens,
            "underscore" => PdfCore.PdfTabLeaderStyle.Underscores,
            _ => PdfCore.PdfTabLeaderStyle.None
        };
        double left, right;
        if (tableDefaults.HasValue) {
            var indentation = ResolveNativeTableCellParagraphIndentation(paragraph, tableDefaults.Value,
                WordDocumentTraversal.GetListInfo(paragraph)?.MarkerVisible == true);
            left = indentation.Left;
            right = indentation.Right;
        } else {
            PdfCore.PdfParagraphStyle style = CreateNativeParagraphStyle(paragraph, defaults);
            if (WordDocumentTraversal.GetListInfo(paragraph)?.MarkerVisible == true) ApplyNativeInlineListIndent(paragraph, style);
            left = style.LeftIndent;
            right = style.RightIndent;
        }
        // Native controls distinguish center tabs from a symmetric midpoint of
        // the indented frame: an indent-relative center follows the leading
        // indent but ignores the trailing indent. Right tabs still stop at it.
        double offset = tab.RelativeTo?.InnerText == "indent"
            ? alignment == "center" ? right / 2D : 0D
            : fraction * (left + right) - left;
        return PdfCore.PdfTextRun.RelativeTab(pdfAlignment, leader, fraction, offset);
    }

    private static bool HasNativePositionalTabs(WordParagraph run) =>
        run.EnumerateEffectiveRunContent().Any(child => child is W.PositionalTab);
}

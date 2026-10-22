using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Translates source paragraph justification policy without duplicating PDF line layout.</summary>
    private static void ApplyNativeParagraphJustification(PdfCore.PdfParagraphStyle style, WordParagraph paragraph,
        W.JustificationValues? resolvedAlignment = null) {
        W.JustificationValues? alignment = resolvedAlignment ?? ResolveNativeParagraphJustification(paragraph);
        if (alignment != W.JustificationValues.Both && alignment != W.JustificationValues.Distribute) return;
        W.DoNotExpandShiftReturn? suppression = paragraph._document?._wordprocessingDocument?.MainDocumentPart?
            .DocumentSettingsPart?.Settings?.GetFirstChild<W.Compatibility>()?.GetFirstChild<W.DoNotExpandShiftReturn>();
        style.JustifyExplicitLineBreaks = ReadNativeOnOff(suppression) != true;
        if (alignment == W.JustificationValues.Distribute) {
            style.JustificationMode = PdfCore.PdfJustificationMode.Characters;
            style.JustifyExplicitLineBreaks = true;
            style.JustifyLastLine = true;
        }
    }
}

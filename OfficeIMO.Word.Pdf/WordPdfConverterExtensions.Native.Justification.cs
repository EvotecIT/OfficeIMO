using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;


public static partial class WordPdfConverterExtensions {
    /// <summary>Translates source paragraph justification policy without duplicating PDF line layout.</summary>
    private static void ApplyNativeParagraphJustification(PdfCore.PdfParagraphStyle style, WordParagraph paragraph,
        string? resolvedAlignment = null) {
        string? alignment = resolvedAlignment ?? ResolveNativeParagraphJustification(paragraph);
        // A right-to-left paragraph ends justified text, which Word does not stretch, at its right edge.
        style.AlignJustifiedRemainderRight = IsNativeBiDiParagraph(paragraph) && alignment != "distribute" &&
            alignment != "thaiDistribute" && MapNativeParagraphAlign(alignment) == global::OfficeIMO.OfficeAlignment.Justify;
        if (alignment != "both" && alignment != "distribute") return;
        style.UseWordJustificationSpacing = true;
        OfficeXmlNode suppression = paragraph._document
            .SettingsNode().WChild("compat").WChild("doNotExpandShiftReturn");
        style.JustifyExplicitLineBreaks = ReadNativeOnOff(suppression) != true;
        if (alignment == "distribute") {
            style.JustificationMode = PdfCore.PdfJustificationMode.Characters;
            style.JustifyExplicitLineBreaks = true;
            style.JustifyLastLine = true;
        }
    }
}

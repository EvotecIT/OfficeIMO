
namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Reports unsupported frame layout without changing paragraph or inherited style XML.</summary>
    private static void RecordNativeParagraphFrameDiagnostics(WordParagraph paragraph, WordPdfExportOptions options,
        string source) {
        OfficeXmlNode node = paragraph.Node();
        if (node.IsNone || !options.DiagnosedFrameParagraphs.Add(node)) return;
        if (HasNativeParagraphFrame(paragraph) && !CanRenderNativeParagraphFrame(paragraph)) AddNativeExportWarning(options, "native-paragraph-frame-flattened", source,
            "Word paragraph frame positioning, dimensions and wrapping are not supported by PDF conversion; its visible content is placed in ordinary flow.");
    }

    private static bool HasNativeParagraphFrame(WordParagraph paragraph) =>
        paragraph.Node().WChild("pPr").WChild("framePr").Exists ||
            GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId)
                .Any(style => style.WChild("pPr").WChild("framePr").Exists) ||
            paragraph._document.DocumentDefaultsNode().WChild("pPrDefault").WChild("pPr").WChild("framePr").Exists;
}

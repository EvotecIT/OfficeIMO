using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Reports unsupported frame layout without changing paragraph or inherited style XML.</summary>
    private static void RecordNativeParagraphFrameDiagnostics(WordParagraph paragraph, WordToPdfOptions options,
        string source) {
        if (paragraph._paragraph == null || !options.DiagnosedFrameParagraphs.Add(paragraph._paragraph)) return;
        bool hasFrame = paragraph._paragraph.ParagraphProperties?.GetFirstChild<W.FrameProperties>() != null ||
            GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId)
                .Any(style => style.StyleParagraphProperties?.GetFirstChild<W.FrameProperties>() != null) ||
            paragraph._document._wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?
                .GetFirstChild<W.ParagraphPropertiesDefault>()?.GetFirstChild<W.ParagraphPropertiesBaseStyle>()?
                .GetFirstChild<W.FrameProperties>() != null;
        if (hasFrame) AddNativeExportWarning(options, "NativeParagraphFrameFlattened", source,
            "Word paragraph frame positioning, dimensions and wrapping are not supported by PDF conversion; its visible content is placed in ordinary flow.");
    }
}

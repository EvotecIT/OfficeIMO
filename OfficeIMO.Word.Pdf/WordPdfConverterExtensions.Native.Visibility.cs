using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static bool GetNativeParagraphHiddenDefault(WordParagraph paragraph) =>
        GetNativeParagraphStyleDefaults(paragraph).Hidden ??
        ReadNativeOnOff(paragraph._document._wordprocessingDocument.MainDocumentPart?
            .StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?
            .RunPropertiesBaseStyle?.GetFirstChild<W.Vanish>()) ?? false;

    private static bool ShouldRenderNativeEmptyParagraphLineBox(WordParagraph paragraph) {
        // A paragraph mark inherits run defaults independently of the visible text.
        bool hiddenMark = ReadNativeOnOff(paragraph._paragraph?.ParagraphProperties?
            .ParagraphMarkRunProperties?.GetFirstChild<W.Vanish>()) ??
            GetNativeParagraphHiddenDefault(paragraph);
        return !hiddenMark;
    }

    private static bool IsNativeHiddenTextRun(WordParagraph paragraph, WordParagraph? fallback = null) {
        W.RunProperties? runProperties = GetNativeRunProperties(paragraph);
        NativeCharacterStyleDefaults characterStyle = GetNativeCharacterStyleDefaults(paragraph._document, runProperties);
        return ReadNativeOnOff(runProperties?.GetFirstChild<W.Vanish>()) ??
            characterStyle.Hidden ?? GetNativeParagraphHiddenDefault(fallback ?? paragraph);
    }
}

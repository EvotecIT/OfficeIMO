namespace OfficeIMO.Word.Pdf;


public static partial class WordPdfConverterExtensions {
    private static bool GetNativeParagraphHiddenDefault(WordParagraph paragraph) =>
        GetNativeParagraphStyleDefaults(paragraph).Hidden ??
        WordStyleCatalog.For(paragraph._document).DocumentDefaultsNode().WChild("rPrDefault").WChild("rPr")
            .WChild("vanish").OnOff() ?? false;

    private static bool ShouldRenderNativeEmptyParagraphLineBox(WordParagraph paragraph) {
        // A paragraph mark inherits run defaults independently of the visible text.
        bool hiddenMark = paragraph.Node().WChild("pPr").WChild("rPr").WChild("vanish").OnOff() ??
            GetNativeParagraphHiddenDefault(paragraph);
        return !hiddenMark;
    }

    private static bool IsNativeHiddenTextRun(WordParagraph paragraph, WordParagraph? fallback = null) {
        OfficeXmlNode runProperties = GetNativeRunProperties(paragraph);
        NativeCharacterStyleDefaults characterStyle = GetNativeCharacterStyleDefaults(paragraph._document, runProperties);
        return runProperties.WChild("vanish").OnOff() ??
            characterStyle.Hidden ?? GetNativeParagraphHiddenDefault(fallback ?? paragraph);
    }
}

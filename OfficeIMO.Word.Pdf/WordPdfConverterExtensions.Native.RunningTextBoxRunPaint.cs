using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Rejects unsupported effects on visible runs; inherited toggle combinations are conservatively approximated.</summary>
    private static bool HasNativeUnsupportedRunningTextBoxRunPaint(WordParagraph paragraph, IReadOnlyList<WordParagraph> runs) {
        IReadOnlyList<W.Style> paragraphStyles = GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId);
        W.RunPropertiesBaseStyle? documentProperties = paragraph._document._wordprocessingDocument.MainDocumentPart?
            .StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        foreach (WordParagraph run in runs) {
            if (IsNativeHiddenTextRun(run, paragraph) || !run.EnumerateEffectiveRunContent().OfType<W.Text>()
                .Any(text => !string.IsNullOrWhiteSpace(text.Text))) continue;
            W.RunProperties? direct = GetNativeRunProperties(run);
            IReadOnlyList<W.Style> characterStyles = GetNativeCharacterStyleChain(paragraph._document,
                direct?.RunStyle?.Val?.Value);
            if (NativeRunningTextBoxRunMayHaveEffect<W.Shadow>(direct, paragraphStyles, characterStyles, documentProperties) ||
                NativeRunningTextBoxRunMayHaveEffect<W.Outline>(direct, paragraphStyles, characterStyles, documentProperties) ||
                NativeRunningTextBoxRunMayHaveEffect<W.Emboss>(direct, paragraphStyles, characterStyles, documentProperties) ||
                NativeRunningTextBoxRunMayHaveEffect<W.Imprint>(direct, paragraphStyles, characterStyles, documentProperties)) return true;
        }
        return false;
    }

    private static bool NativeRunningTextBoxRunMayHaveEffect<T>(W.RunProperties? direct,
        IReadOnlyList<W.Style> paragraphStyles, IReadOnlyList<W.Style> characterStyles,
        W.RunPropertiesBaseStyle? documentProperties) where T : W.OnOffType {
        bool? declared = ReadNativeOnOff(direct?.GetFirstChild<T>());
        if (declared.HasValue) return declared.Value;
        // A direct false is absolute. At each style level the nearest declaration wins.
        // Enabled values at multiple levels can toggle each other; this narrow admission
        // profile keeps that ambiguous unsupported paint explicitly approximated.
        return ReadNativeNearestRunningTextBoxStyleEffect<T>(paragraphStyles) == true ||
            ReadNativeNearestRunningTextBoxStyleEffect<T>(characterStyles) == true ||
            ReadNativeOnOff(documentProperties?.GetFirstChild<T>()) == true;
    }

    private static bool? ReadNativeNearestRunningTextBoxStyleEffect<T>(IReadOnlyList<W.Style> styles) where T : W.OnOffType {
        for (int index = styles.Count - 1; index >= 0; index--) {
            bool? declared = ReadNativeOnOff(styles[index].GetFirstChild<W.StyleRunProperties>()?.GetFirstChild<T>());
            if (declared.HasValue) return declared.Value;
        }
        return null;
    }
}

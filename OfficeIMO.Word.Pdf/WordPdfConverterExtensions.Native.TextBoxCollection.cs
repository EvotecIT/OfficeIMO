using System.Collections.Generic;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Uses the same effective visibility owner before discovering any textbox payload.</summary>
    private static List<WordTextBox> GetNativeVisibleTextBoxes(WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs) => runs.Where(run => !IsNativeHiddenTextRun(run, paragraph))
        .SelectMany(run => run.GetTextBoxes()).ToList();
}

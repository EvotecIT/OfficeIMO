using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Uses the same effective visibility owner before discovering any textbox payload.</summary>
    private static List<WordTextBox> GetNativeVisibleTextBoxes(WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs) {
        // Header and footer building blocks (grouped page numbers) keep their text in a text box
        // inside a shape group. Their text is what matters there, so read it like a plain text box.
        bool inHeaderFooter = paragraph._paragraph.Ancestors().Any(parent => parent is W.Header or W.Footer);
        return runs.Where(run => !IsNativeHiddenTextRun(run, paragraph))
            .SelectMany(run => inHeaderFooter ? run.GetTextBoxes().Concat(run.GetGroupTextBoxes()) : run.GetTextBoxes())
            .ToList();
    }

    private static void RecordNativeCellTextBoxDiagnostic(WordParagraph paragraph,
        WordPdfExportOptions options, string source) {
        if (!paragraph._paragraph.Ancestors<W.TableCell>().Any() ||
            paragraph._paragraph.Ancestors().Any(parent => parent is W.Footnote or W.Endnote) ||
            IsNativeBarInMergeContinuation(paragraph)) return;
        List<WordParagraph> runs = GetNativeRuns(paragraph, includeInlineContentControls: true);
        if (GetNativeVisibleTextBoxes(paragraph, runs).Count == 0 ||
            !options.DiagnosedCellTextBoxParagraphs.Add(paragraph._paragraph)) return;
        AddNativeExportWarning(options, "native-mixed-text-box-layout-approximated", source + " text box",
            "Table-cell text boxes use flattened inline text. Inner paragraph boundaries and formatting, anchored positioning, shape geometry, wrapping and pagination are approximated.");
    }
}

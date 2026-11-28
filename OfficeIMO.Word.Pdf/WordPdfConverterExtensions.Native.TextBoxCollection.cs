
namespace OfficeIMO.Word.Pdf;


public static partial class WordPdfConverterExtensions {
    /// <summary>Uses the same effective visibility owner before discovering any textbox payload.</summary>
    private static List<WordTextBox> GetNativeVisibleTextBoxes(WordParagraph paragraph,
        IReadOnlyList<WordParagraph> runs) {
        // Header and footer building blocks (grouped page numbers) keep their text in a text box
        // inside a shape group. Their text is what matters there, so read it like a plain text box.
        bool inHeaderFooter = paragraph.Node().Ancestors().Any(parent => parent.IsW("hdr") || parent.IsW("ftr"));
        return runs.Where(run => !IsNativeHiddenTextRun(run, paragraph))
            .SelectMany(run => inHeaderFooter ? run.GetTextBoxes().Concat(run.GetGroupTextBoxes()) : run.GetTextBoxes())
            .ToList();
    }

    private static void RecordNativeCellTextBoxDiagnostic(WordParagraph paragraph,
        WordPdfExportOptions options, string source) {
        OfficeXmlNode node = paragraph.Node();
        if (node.Ancestor(WordXml.W, "tc").IsNone ||
            node.Ancestors().Any(parent => parent.IsW("footnote") || parent.IsW("endnote")) ||
            IsNativeBarInMergeContinuation(paragraph)) return;
        List<WordParagraph> runs = GetNativeRuns(paragraph, includeInlineContentControls: true);
        if (GetNativeVisibleTextBoxes(paragraph, runs).Count == 0 ||
            !options.DiagnosedCellTextBoxParagraphs.Add(node)) return;
        AddNativeExportWarning(options, "native-mixed-text-box-layout-approximated", source + " text box",
            "Table-cell text boxes use flattened inline text. Inner paragraph boundaries and formatting, anchored positioning, shape geometry, wrapping and pagination are approximated.");
    }
}

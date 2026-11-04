using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Word;

public partial class WordSection {
    /// <summary>
    /// Wraps existing blocks in source order without inserting content or initializing
    /// table children. Paragraphs use the same run-level projection as Elements;
    /// pure section boundaries are structural and have no returned Word element.
    /// </summary>
    internal static IReadOnlyList<WordElement> WrapBlockElements(
        WordDocument document,
        IEnumerable<OpenXmlElement> blocks,
        CancellationToken cancellationToken = default) {
        var elements = new List<WordElement>();
        foreach (OpenXmlElement block in blocks) {
            cancellationToken.ThrowIfCancellationRequested();
            switch (block) {
                case Paragraph paragraph when !IsPureSectionBreakParagraph(paragraph):
                    elements.AddRange(ConvertParagraphToWordParagraphs(document, paragraph,
                        cancellationToken: cancellationToken));
                    break;
                case Table table:
                    elements.Add(new WordTable(document, table, initializeChildren: false));
                    break;
                case SdtBlock control:
                    // Suppress conversion-time field updates; the returned model remains editable.
                    elements.Add(ConvertStdBlockToWordElements(document, control, readOnly: true));
                    break;
            }
        }
        return elements.AsReadOnly();
    }
}

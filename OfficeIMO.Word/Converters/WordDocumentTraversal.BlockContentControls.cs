using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading;

namespace OfficeIMO.Word;

internal static partial class WordDocumentTraversal {
    private const int MaximumBlockContentControlDepth = 128;

    /// <summary>Projects the immediate cached children of a Word block content control.</summary>
    internal static IReadOnlyList<WordElement> GetBlockContentElements(WordDocument document, SdtBlock? block,
        Func<Paragraph, bool>? includeParagraph = null, CancellationToken cancellationToken = default) {
        var elements = new List<WordElement>();
        if (block?.SdtContentBlock is not { } content) {
            return elements;
        }

        foreach (OpenXmlElement child in content.ChildElements) {
            cancellationToken.ThrowIfCancellationRequested();
            if (child is Paragraph paragraph && (includeParagraph == null || includeParagraph(paragraph))) {
                elements.AddRange(WordSection.ConvertParagraphToWordParagraphs(document, paragraph, cancellationToken: cancellationToken));
            } else if (child is Table table) {
                elements.Add(new WordTable(document, table, initializeChildren: false, readOnly: true));
            } else if (child is SdtBlock nested) {
                elements.Add(WordSection.ConvertStdBlockToWordElements(document, nested, readOnly: true));
            }
        }

        return elements;
    }

    /// <summary>Expands ordinary block controls in place for formats with only static block content.</summary>
    internal static List<WordElement> ExpandBlockContentControls(List<WordElement> elements, CancellationToken cancellationToken) {
        if (!elements.Any(element => element is WordStructuredDocumentTag { SdtBlock: not null })) {
            return elements;
        }

        var projected = new List<WordElement>(elements.Count);
        Append(elements, depth: 0);
        return projected;

        void Append(IReadOnlyList<WordElement> source, int depth) {
            foreach (WordElement element in source) {
                cancellationToken.ThrowIfCancellationRequested();
                (WordDocument Document, SdtBlock Block)? blockContent = element switch {
                    WordStructuredDocumentTag { SdtBlock: { } block } control => (control.Document, block),
                    WordCoverPage coverPage when depth > 0 => (coverPage.Document, coverPage.SdtBlock),
                    _ => null
                };
                if (blockContent is not { } content) {
                    projected.Add(element);
                    continue;
                }

                if (depth >= MaximumBlockContentControlDepth) {
                    throw new InvalidDataException(
                        $"Structured document tag nesting exceeds the supported limit of {MaximumBlockContentControlDepth} levels.");
                }

                Append(GetBlockContentElements(content.Document, content.Block, cancellationToken: cancellationToken), depth + 1);
            }
        }
    }
}

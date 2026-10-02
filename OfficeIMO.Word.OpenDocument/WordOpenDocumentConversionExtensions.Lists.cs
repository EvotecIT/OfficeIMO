using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using System.Xml.Linq;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentConversionExtensions {
    private sealed class OdtListConversionState {
        private readonly Dictionary<XElement, WordList> _lists = new();
        private readonly HashSet<XElement> _items = new();

        internal WordParagraph AddParagraph(OdtContentBlock block, Func<bool, WordList> createList,
            Func<WordParagraph> createParagraph, Action<WordParagraph> appendItem,
            OdfConversionReport report, out bool createdList) {
            XElement source = block.Paragraph!.Element;
            XElement list = source.Ancestors(OdfNamespaces.Text + "list").First();
            // A list header has no item in its immediate containing list. A
            // continuation paragraph belongs to the already numbered item.
            XElement? item = source.Ancestors().TakeWhile(element => element != list)
                .FirstOrDefault(element => element.Name == OdfNamespaces.Text + "list-item");
            createdList = false;
            if (item == null) return createParagraph();
            if (!_lists.TryGetValue(list, out WordList? targetList)) {
                targetList = createList(block.IsOrderedList == true);
                _lists.Add(list, targetList);
                createdList = true;
                if (list.Attribute(OdfNamespaces.Text + "continue-list") != null ||
                    (string?)list.Attribute(OdfNamespaces.Text + "continue-numbering") is "true" or "1") {
                    report.Add("list-numbering", OdfConversionMappingStatus.Approximated, 1,
                        "Explicit continuation between separate ODT lists is not retained.");
                }
            }
            int level = Math.Max(0, Math.Min(8, block.ListLevel));
            if (block.ListLevel > 8) report.Add("list-levels", OdfConversionMappingStatus.Approximated, 1);
            if (!_items.Add(item)) {
                WordParagraph continuation = createParagraph();
                WordListLevel? definition = targetList.Numbering.Levels
                    .FirstOrDefault(value => value._level.LevelIndex?.Value == level);
                if (definition != null) {
                    continuation.IndentationBefore = definition.IndentationLeft;
                    continuation.IndentationFirstLine = 0;
                }
                return continuation;
            }
            if (item.Attribute(OdfNamespaces.Text + "start-value") != null) {
                report.Add("list-numbering", OdfConversionMappingStatus.Approximated, 1,
                    "Per-item ODT numbering restarts are not retained.");
            }
            WordParagraph paragraph = targetList.AddItem(null, level);
            // AddItem normally follows the prior numbered item, which may precede
            // continuation paragraphs or an intervening nested table/list.
            appendItem(paragraph);
            return paragraph;
        }
    }
}

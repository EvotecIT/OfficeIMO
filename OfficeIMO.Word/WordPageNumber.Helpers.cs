using System.Collections.Concurrent;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

public partial class WordPageNumber {
    private const int W = OfficeXmlNamespaces.WordprocessingML;
    private static readonly ConcurrentDictionary<WordPageNumberStyle, Lazy<OfficeXmlNode>> Templates = new();

    /// <summary>Appends text to the last paragraph of the page number.</summary>
    public WordParagraph AppendText(string text) {
        if (string.IsNullOrEmpty(text)) throw new ArgumentNullException(nameof(text));
        return _listParagraphs[_listParagraphs.Count - 1].AddText(text);
    }

    /// <summary>Appends a page-number style to the specified header.</summary>
    public WordPageNumber(WordDocument wordDocument, WordHeader wordHeader, WordPageNumberStyle wordPageNumberStyle)
        : this(wordDocument, (WordHeaderFooter)wordHeader, wordPageNumberStyle) {
    }

    /// <summary>Appends a page-number style to the specified footer.</summary>
    public WordPageNumber(WordDocument wordDocument, WordFooter wordFooter, WordPageNumberStyle wordPageNumberStyle)
        : this(wordDocument, (WordHeaderFooter)wordFooter, wordPageNumberStyle) {
    }

    private WordPageNumber(WordDocument document, WordHeaderFooter story, WordPageNumberStyle style) {
        string name = style.ToString();
        if (ReferenceEquals(document, story.Document) && document._ownedModel is { } model &&
            story.StoryRoot is { Tree: not null } root) {
            OfficeXmlNode template = Templates.GetOrAdd(style, _ => new Lazy<OfficeXmlNode>(() => {
                byte[] markup = WordEmbeddedMarkup.GetBuildingBlockMarkup(WordEmbeddedMarkup.PageNumbers, name, nameof(style));
                OfficeXmlTree tree = OfficeXmlTree.Parse(markup, markup.Length, new OfficeXmlNamespaces.Table());
                return new OfficeXmlNode(tree, tree.Root);
            })).Value;
            if (CanInsertOwnedTemplate(model, template)) {
                OfficeXmlTree storyTree = root.Tree!;
                var block = new OfficeXmlNode(storyTree, storyTree.ImportElement(template.Tree!, template.Node));
                document.AssignNewSdtIds(block);
                WordDrawingIdAllocator.ReassignOwned(model, block);
                DeclareGalleryNamespaces(block);
                root.Append(block);
                story.TrackOwnedParagraphs(root.Tree!);
                model.MarkChanged();
                _listParagraphs = WordOwnedParagraphs.ProjectFlatFieldParagraphs(document, block.Child(W, "sdtContent"));
                _wordParagraph = _listParagraphs[0];
                return;
            }
        }

        _listParagraphs = story.AppendSdkPageNumberTemplate(document, name);
        _wordParagraph = _listParagraphs[0];
    }

    private static bool CanInsertOwnedTemplate(WordPackageModel model, OfficeXmlNode template) {
        foreach (OfficeXmlNode node in WordFieldInventory.EnumerateFieldContent(template)) {
            if (node.Is(W, "sdt") || node.Is(W, "hyperlink")) return false;
            if ((node.Is(W, "pStyle") || node.Is(W, "rStyle") || node.Is(W, "tblStyle")) &&
                node.Attribute(W, "val") is { } style && !model.Styles.DefinesWithDependencies(style)) return false;
        }
        return true;
    }

    internal static void DeclareGalleryNamespaces(OfficeXmlNode block) {
        block.DeclareNamespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        block.DeclareNamespace("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
        block.DeclareNamespace("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
        block.DeclareNamespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
    }
}

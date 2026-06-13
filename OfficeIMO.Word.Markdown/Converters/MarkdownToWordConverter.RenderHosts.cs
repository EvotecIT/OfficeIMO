using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Html;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown {
    internal partial class MarkdownToWordConverter {
        private interface IWordBlockRenderHost {
            WordParagraph CreateParagraph();
            WordList CreateList(WordListStyle style);
            WordTable CreateTable(int rows, int columns);
            bool TryAddTableOfContents(int minLevel, int maxLevel, string? title);
            bool SupportsHtmlInsertion { get; }
            void InsertHtml(string html);
            bool SupportsHorizontalRule { get; }
            void InsertHorizontalRule();
            void NotifyListRendered(WordList list);
        }

        private sealed class DocumentWordBlockRenderHost : IWordBlockRenderHost {
            private readonly WordDocument _document;

            public DocumentWordBlockRenderHost(WordDocument document) {
                _document = document ?? throw new ArgumentNullException(nameof(document));
            }

            public WordParagraph CreateParagraph() => _document.AddParagraph(string.Empty);
            public WordList CreateList(WordListStyle style) => _document.AddList(style);
            public WordTable CreateTable(int rows, int columns) => _document.AddTable(rows, columns);
            public bool TryAddTableOfContents(int minLevel, int maxLevel, string? title) {
                var toc = _document.AddTableOfContent(minLevel: minLevel, maxLevel: maxLevel);
                toc.Text = string.IsNullOrWhiteSpace(title) ? string.Empty : title!;

                return true;
            }

            public bool SupportsHtmlInsertion => true;
            public void InsertHtml(string html) => _document.AddHtmlToBody(html);
            public bool SupportsHorizontalRule => true;
            public void InsertHorizontalRule() => _document.AddHorizontalLine();
            public void NotifyListRendered(WordList list) { }
        }

        private sealed class TableCellWordBlockRenderHost : IWordBlockRenderHost {
            private readonly WordTableCell _cell;
            private bool _wroteContent;

            public TableCellWordBlockRenderHost(WordTableCell cell) {
                _cell = cell ?? throw new ArgumentNullException(nameof(cell));
            }

            public WordParagraph CreateParagraph() {
                if (!_wroteContent) {
                    var existing = _cell.Paragraphs.FirstOrDefault();
                    if (existing != null) {
                        _wroteContent = true;
                        return existing;
                    }
                }

                _wroteContent = true;
                return _cell.AddParagraph();
            }

            public WordList CreateList(WordListStyle style) {
                _wroteContent = true;
                return _cell.AddList(style);
            }

            public WordTable CreateTable(int rows, int columns) {
                _wroteContent = true;
                return _cell.AddTable(rows, columns);
            }

            public bool TryAddTableOfContents(int minLevel, int maxLevel, string? title) => false;

            public bool SupportsHtmlInsertion => false;
            public void InsertHtml(string html) { }
            public bool SupportsHorizontalRule => false;
            public void InsertHorizontalRule() { }
            public void NotifyListRendered(WordList list) { }
        }

        private sealed class HeaderFooterWordBlockRenderHost : IWordBlockRenderHost {
            private readonly WordHeaderFooter _headerFooter;

            public HeaderFooterWordBlockRenderHost(WordHeaderFooter headerFooter) {
                _headerFooter = headerFooter ?? throw new ArgumentNullException(nameof(headerFooter));
            }

            public WordParagraph CreateParagraph() => _headerFooter.AddParagraph(string.Empty);
            public WordList CreateList(WordListStyle style) => _headerFooter.AddList(style);
            public WordTable CreateTable(int rows, int columns) => _headerFooter.AddTable(rows, columns);
            public bool TryAddTableOfContents(int minLevel, int maxLevel, string? title) => false;
            public bool SupportsHtmlInsertion => false;
            public void InsertHtml(string html) { }
            public bool SupportsHorizontalRule => true;
            public void InsertHorizontalRule() => _headerFooter.AddHorizontalLine();
            public void NotifyListRendered(WordList list) { }
        }

        private sealed class BodyInsertionPointWordBlockRenderHost : IWordBlockRenderHost {
            private readonly WordDocument _document;
            private readonly OpenXmlElement _anchor;
            private readonly List<Paragraph> _listSeeds = new();

            public BodyInsertionPointWordBlockRenderHost(WordDocument document, OpenXmlElement anchor) {
                _document = document ?? throw new ArgumentNullException(nameof(document));
                _anchor = anchor ?? throw new ArgumentNullException(nameof(anchor));
                if (_anchor.Parent == null) {
                    throw new InvalidOperationException("Insertion anchor must be attached to a Word document.");
                }
            }

            public WordParagraph CreateParagraph() {
                var paragraph = new Paragraph();
                _anchor.InsertBeforeSelf(paragraph);
                return new WordParagraph(_document, paragraph);
            }

            public WordList CreateList(WordListStyle style) {
                var seed = CreateSeedParagraph();
                _listSeeds.Add(seed._paragraph);
                return seed.AddList(style);
            }

            public WordTable CreateTable(int rows, int columns) {
                var seed = CreateSeedParagraph();
                try {
                    return seed.AddTableAfter(rows, columns);
                } finally {
                    seed.Remove();
                }
            }

            public bool TryAddTableOfContents(int minLevel, int maxLevel, string? title) => false;

            public bool SupportsHtmlInsertion => false;
            public void InsertHtml(string html) { }
            public bool SupportsHorizontalRule => true;
            public void InsertHorizontalRule() => CreateParagraph().AddHorizontalLine();

            public void NotifyListRendered(WordList list) {
                foreach (var seed in _listSeeds) {
                    seed.Remove();
                }
                _listSeeds.Clear();
            }

            private WordParagraph CreateSeedParagraph() {
                var paragraph = new Paragraph();
                _anchor.InsertBeforeSelf(paragraph);
                return new WordParagraph(_document, paragraph);
            }
        }
    }
}

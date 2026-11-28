using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Drawing;
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
            private readonly MarkdownWordExportOptions _options;

            public DocumentWordBlockRenderHost(WordDocument document, MarkdownWordExportOptions options) {
                _options = options;
                _document = document ?? throw new ArgumentNullException(nameof(document));
            }

            public WordParagraph CreateParagraph() => _document.AddParagraph(string.Empty);
            public WordList CreateList(WordListStyle style) => _document.AddList(style);
            public WordTable CreateTable(int rows, int columns) => _document.AddTable(rows, columns, WordTableStyle.TableNormal);
            public bool TryAddTableOfContents(int minLevel, int maxLevel, string? title) {
                var toc = _document.AddTableOfContent(minLevel: minLevel, maxLevel: maxLevel);
                toc.Text = string.IsNullOrWhiteSpace(title) ? string.Empty : title!;

                return true;
            }

            public bool SupportsHtmlInsertion => true;
            public void InsertHtml(string html) {
                if (_options.RawHtmlRenderer?.Invoke(html, _document, _options.DiagnosticSink) == true) return;
                _document.AddParagraph(html);
                _options.DiagnosticSink?.Invoke(new WordMarkdownConversionDiagnostic(
                    OfficeDiagnosticCodes.HtmlBlockAsText,
                    "Raw HTML was retained as literal text because no HTML renderer was configured.",
                    OfficeConversionLossKind.Approximation));
            }
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
                return _cell.AddTable(rows, columns, WordTableStyle.TableNormal);
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
            public WordTable CreateTable(int rows, int columns) => _headerFooter.AddTable(rows, columns, WordTableStyle.TableNormal);
            public bool TryAddTableOfContents(int minLevel, int maxLevel, string? title) => false;
            public bool SupportsHtmlInsertion => false;
            public void InsertHtml(string html) { }
            public bool SupportsHorizontalRule => true;
            public void InsertHorizontalRule() => _headerFooter.AddHorizontalLine();
            public void NotifyListRendered(WordList list) { }
        }

        private sealed class BodyInsertionPointWordBlockRenderHost : IWordBlockRenderHost {
            private readonly WordDocument _document;
            private readonly OfficeXmlNode _anchor;
            private readonly List<OfficeXmlNode> _listSeeds = new();

            public BodyInsertionPointWordBlockRenderHost(WordDocument document, OfficeXmlNode anchor) {
                _document = document ?? throw new ArgumentNullException(nameof(document));
                if (anchor.IsNone) throw new ArgumentNullException(nameof(anchor));
                _anchor = anchor;
                if (_anchor.Parent.IsNone) {
                    throw new InvalidOperationException("Insertion anchor must be attached to a Word document.");
                }
            }

            public WordParagraph CreateParagraph() => _document.ParagraphView(InsertParagraphBeforeAnchor());

            public WordList CreateList(WordListStyle style) {
                var seed = CreateSeedParagraph();
                _listSeeds.Add(seed.Node());
                return seed.AddList(style);
            }

            public WordTable CreateTable(int rows, int columns) {
                var seed = CreateSeedParagraph();
                try {
                    return seed.AddTableAfter(rows, columns, WordTableStyle.TableNormal);
                } finally {
                    seed.Remove();
                }
            }

            public bool TryAddTableOfContents(int minLevel, int maxLevel, string? title) {
                var toc = _document.AddTableOfContent(minLevel: minLevel, maxLevel: maxLevel);
                toc.Text = string.IsNullOrWhiteSpace(title) ? string.Empty : title!;

                OfficeXmlNode block = toc.SdtBlockNode();
                block.Remove();
                block = _anchor.Parent.InsertBefore(block, _anchor);

                OfficeXmlNode body = BodyRoot(_document);
                _document._tableOfContentIndex = body.Elements().ToList().IndexOf(block);
                return true;
            }

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

            private WordParagraph CreateSeedParagraph() => _document.ParagraphView(InsertParagraphBeforeAnchor());

            private OfficeXmlNode InsertParagraphBeforeAnchor() {
                OfficeXmlNode parent = _anchor.Parent;
                return parent.InsertBefore(parent.CreateElement(WordXml.W, "p"), _anchor);
            }
        }
    }
}

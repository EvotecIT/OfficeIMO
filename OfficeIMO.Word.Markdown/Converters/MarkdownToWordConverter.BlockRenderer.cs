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
        private sealed class BlockRenderer : Omd.MarkdownVisitor {
            private readonly IWordBlockRenderHost _host;
            private readonly MarkdownToWordOptions _options;
            private readonly WordDocument _document;
            private readonly int _listLevel;
            private readonly int _quoteDepth;
            private readonly double _pageContentWidthPixels;
            private readonly Omd.ColumnAlignment _alignment;

            public BlockRenderer(
                IWordBlockRenderHost host,
                MarkdownToWordOptions options,
                WordDocument document,
                int listLevel,
                int quoteDepth,
                double pageContentWidthPixels,
                Omd.ColumnAlignment alignment) {
                _host = host ?? throw new ArgumentNullException(nameof(host));
                _options = options ?? throw new ArgumentNullException(nameof(options));
                _document = document ?? throw new ArgumentNullException(nameof(document));
                _listLevel = listLevel;
                _quoteDepth = quoteDepth;
                _pageContentWidthPixels = pageContentWidthPixels;
                _alignment = alignment;
            }

            public void Render(Omd.IMarkdownBlock block) {
                if (block == null) {
                    return;
                }

                if (block is Omd.MarkdownObject markdownObject) {
                    Visit(markdownObject);
                } else {
                    RenderFallback(block);
                }
            }

            private void RenderNested(
                Omd.IMarkdownBlock block,
                int? listLevel = null,
                int? quoteDepth = null,
                double? pageContentWidthPixels = null,
                Omd.ColumnAlignment? alignment = null) {
                new BlockRenderer(
                    _host,
                    _options,
                    _document,
                    listLevel ?? _listLevel,
                    quoteDepth ?? _quoteDepth,
                    pageContentWidthPixels ?? _pageContentWidthPixels,
                    alignment ?? _alignment)
                    .Render(block);
            }

            private void RenderFallback(Omd.IMarkdownBlock block) {
                var fallback = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(fallback, _quoteDepth, _alignment);
                fallback.AddText(block.RenderMarkdown());
            }

            protected override void VisitBlock(Omd.MarkdownBlock block) {
                if (block is Omd.IMarkdownBlock markdownBlock) {
                    RenderFallback(markdownBlock);
                }
            }

            protected override void VisitHeadingBlock(Omd.HeadingBlock block) {
                var headingParagraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(headingParagraph, _quoteDepth, _alignment);
                ProcessInlinesOmd(block.Inlines, headingParagraph, _options, _document, _currentFootnotes, _pageContentWidthPixels, _listLevel, _quoteDepth);
                headingParagraph.Style = HeadingStyleMapper.GetHeadingStyleForLevel(block.Level);
            }

            protected override void VisitParagraphBlock(Omd.ParagraphBlock block) {
                var paragraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                ProcessInlinesOmd(block.Inlines, paragraph, _options, _document, _currentFootnotes, _pageContentWidthPixels, _listLevel, _quoteDepth);
            }

            protected override void VisitImageBlock(Omd.ImageBlock block) {
                var paragraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                RenderMarkdownImageIntoParagraph(
                    paragraph,
                    block.Path ?? string.Empty,
                    block.Alt,
                    block.Width,
                    block.Height,
                    _options,
                    _pageContentWidthPixels,
                    _listLevel,
                    _quoteDepth,
                    "block");

                if (!string.IsNullOrWhiteSpace(block.Caption)) {
                    var captionParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(captionParagraph, _quoteDepth, _alignment);
                    captionParagraph.AddText(block.Caption!);
                }
            }

            protected override void VisitCodeBlock(Omd.CodeBlock block) {
                var codeParagraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(codeParagraph, _quoteDepth, _alignment);
                var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";
                codeParagraph.AddFormattedText(block.Content ?? string.Empty).SetFontFamily(monoFont);
                if (!string.IsNullOrWhiteSpace(block.Caption)) {
                    var captionParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(captionParagraph, _quoteDepth, _alignment);
                    captionParagraph.AddText(block.Caption!);
                }
            }

            protected override void VisitSemanticFencedBlock(Omd.SemanticFencedBlock block) {
                if (TryRenderWordPageBreakSemanticBlock(block, _host, _quoteDepth, _alignment)) {
                    return;
                }

                if (TryRenderWordHeaderFooterSemanticBlock(block, _host, _options, _document, _pageContentWidthPixels)) {
                    return;
                }

                var paragraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";
                paragraph.AddFormattedText(block.Content ?? string.Empty).SetFontFamily(monoFont);
                if (!string.IsNullOrWhiteSpace(block.Caption)) {
                    var captionParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(captionParagraph, _quoteDepth, _alignment);
                    captionParagraph.AddText(block.Caption!);
                }
            }

            protected override void VisitTableBlock(Omd.TableBlock block) =>
                RenderSharedTableBlockOmd(block, _host, _options, _document, _pageContentWidthPixels);

            protected override void VisitUnorderedListBlock(Omd.UnorderedListBlock block) =>
                RenderListBlock(block.Items, WordListStyle.Bulleted, startNumber: null);

            protected override void VisitOrderedListBlock(Omd.OrderedListBlock block) =>
                RenderListBlock(block.Items, WordListStyle.Numbered, block.Start);

            protected override void VisitTocBlock(Omd.TocBlock block) {
                if (block.Scope != Omd.TocScope.Document) {
                    if (block.Entries.Count > 0) {
                        RenderFallback(block);
                    }

                    return;
                }

                int minLevel = NormalizeTocLevel(block.MinLevel, Omd.TocOptions.DefaultMinLevel);
                int maxLevel = NormalizeTocLevel(block.MaxLevel, Omd.TocOptions.DefaultMaxLevel);
                if (block.RequireTopLevel && minLevel > Omd.TocOptions.DefaultMinLevel) {
                    minLevel = Omd.TocOptions.DefaultMinLevel;
                }

                if (maxLevel < minLevel) {
                    maxLevel = minLevel;
                }

                string? title = block.IncludeTitle && !string.IsNullOrWhiteSpace(block.Title)
                    ? block.Title.Trim()
                    : null;

                if (_host.TryAddTableOfContents(minLevel, maxLevel, title)) {
                    return;
                }

                if (block.Entries.Count > 0) {
                    RenderFallback(block);
                }
            }

            protected override void VisitTocMarkerBlock(Omd.TocMarkerBlock block) {
                int minLevel = NormalizeTocLevel(block.MinLevel, Omd.TocOptions.DefaultMinLevel);
                int maxLevel = NormalizeTocLevel(block.MaxLevel, Omd.TocOptions.DefaultMaxLevel);
                if (maxLevel < minLevel) {
                    maxLevel = minLevel;
                }

                string? title = block.IncludeTitle && !string.IsNullOrWhiteSpace(block.Title)
                    ? block.Title.Trim()
                    : null;

                if (!_host.TryAddTableOfContents(minLevel, maxLevel, title)) {
                    RenderFallback(block);
                }
            }

            protected override void VisitHtmlCommentBlock(Omd.HtmlCommentBlock block) {
                if (_host.SupportsHtmlInsertion) {
                    _host.InsertHtml(block.Comment);
                }
            }

            protected override void VisitHtmlRawBlock(Omd.HtmlRawBlock block) {
                if (_host.SupportsHtmlInsertion) {
                    _host.InsertHtml(block.Html);
                } else if (TryRenderHtmlFallbackViaMarkdownAst(block.Html, _host, _options, _document, _quoteDepth, _pageContentWidthPixels, _alignment)) {
                    return;
                } else {
                    var htmlParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(htmlParagraph, _quoteDepth, _alignment);
                    htmlParagraph.AddText(((Omd.IMarkdownBlock)block).RenderMarkdown());
                }
            }

            protected override void VisitHorizontalRuleBlock(Omd.HorizontalRuleBlock block) {
                if (_host.SupportsHorizontalRule) {
                    _host.InsertHorizontalRule();
                } else {
                    var ruleParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(ruleParagraph, _quoteDepth, _alignment);
                    ruleParagraph.AddText("---");
                }
            }

            protected override void VisitDefinitionListBlock(Omd.DefinitionListBlock block) {
                foreach (var entry in block.Entries) {
                    if (entry == null) {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(entry.TermMarkdown) && entry.DefinitionBlocks.Count == 0) {
                        continue;
                    }

                    RenderSharedDefinitionListEntryOmd(entry, _host, _options, _document, _quoteDepth, _pageContentWidthPixels, _alignment);
                }
            }

            protected override void VisitQuoteBlock(Omd.QuoteBlock block) {
                foreach (var child in block.Children) {
                    RenderNested(child, quoteDepth: _quoteDepth + 1);
                }
            }

            protected override void VisitCalloutBlock(Omd.CalloutBlock block) =>
                RenderSharedCalloutBlockOmd(block, _host, _options, _document, _quoteDepth, _pageContentWidthPixels, _alignment);

            protected override void VisitFootnoteDefinitionBlock(Omd.FootnoteDefinitionBlock block) { }

            protected override void VisitDetailsBlock(Omd.DetailsBlock block) {
                if (block.Summary != null) {
                    var summaryParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(summaryParagraph, _quoteDepth, _alignment);
                    ProcessInlinesOmd(block.Summary.Inlines, summaryParagraph, _options, _document, _currentFootnotes, _pageContentWidthPixels, _listLevel, _quoteDepth);
                    foreach (var run in summaryParagraph.GetRuns()) {
                        run.SetBold();
                    }
                }

                foreach (var child in block.ChildBlocks) {
                    RenderNested(child, quoteDepth: _quoteDepth + 1);
                }
            }

            protected override void VisitSummaryBlock(Omd.SummaryBlock block) {
                var summaryParagraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(summaryParagraph, _quoteDepth, _alignment);
                ProcessInlinesOmd(block.Inlines, summaryParagraph, _options, _document, _currentFootnotes, _pageContentWidthPixels, _listLevel, _quoteDepth);
                foreach (var run in summaryParagraph.GetRuns()) {
                    run.SetBold();
                }
            }

            protected override void VisitFrontMatterBlock(Omd.FrontMatterBlock block) {
                if (!_options.RenderFrontMatter) {
                    return;
                }

                var lines = block.Render().Replace("\r", string.Empty).Split('\n');
                var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";

                for (int i = 0; i < lines.Length; i++) {
                    var paragraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                    paragraph.AddFormattedText(lines[i]).SetFontFamily(monoFont);
                }
            }

            private static int NormalizeTocLevel(int level, int fallback) {
                int normalized = level <= 0 ? fallback : level;
                if (normalized < 1) {
                    return 1;
                }

                return normalized > 9 ? 9 : normalized;
            }

            private void RenderListBlock(IReadOnlyList<Omd.ListItem> items, WordListStyle style, int? startNumber) {
                var list = _host.CreateList(style);
                if (startNumber.HasValue && startNumber.Value != 1) {
                    list.Numbering.Levels[0].SetStartNumberingValue(startNumber.Value);
                }

                foreach (var item in items) {
                    var effectiveLevel = _listLevel + item.Level;
                    var firstParagraph = true;
                    var blockChildren = item.BlockChildren;

                    for (int i = 0; i < blockChildren.Count; i++) {
                        if (blockChildren[i] is Omd.ParagraphBlock paragraph) {
                            var listItemParagraph = list.AddItem((string?)null, effectiveLevel);
                            if (firstParagraph && item.IsTask) {
                                listItemParagraph.AddCheckBox(item.Checked);
                            }

                            ApplyBlockParagraphFormatting(listItemParagraph, _quoteDepth, _alignment);
                            ProcessInlinesOmd(paragraph.Inlines, listItemParagraph, _options, _document, _currentFootnotes, _pageContentWidthPixels, effectiveLevel, _quoteDepth);
                            firstParagraph = false;
                            continue;
                        }

                        RenderNested(blockChildren[i], listLevel: effectiveLevel + 1);
                    }
                }

                _host.NotifyListRendered(list);
            }
        }
    }
}

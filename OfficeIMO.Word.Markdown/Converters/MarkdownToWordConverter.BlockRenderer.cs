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
            private readonly MarkdownToWordConverter _converter;
            private readonly IWordBlockRenderHost _host;
            private readonly MarkdownToWordOptions _options;
            private readonly WordDocument _document;
            private readonly int _listLevel;
            private readonly int _quoteDepth;
            private readonly double _pageContentWidthPixels;
            private readonly Omd.ColumnAlignment _alignment;

            public BlockRenderer(
                MarkdownToWordConverter converter,
                IWordBlockRenderHost host,
                MarkdownToWordOptions options,
                WordDocument document,
                int listLevel,
                int quoteDepth,
                double pageContentWidthPixels,
                Omd.ColumnAlignment alignment) {
                _converter = converter ?? throw new ArgumentNullException(nameof(converter));
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
                    _converter,
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
                ApplyBodyTextTheme(fallback, _options);
            }

            protected override void VisitBlock(Omd.MarkdownBlock block) {
                if (block is Omd.IMarkdownBlock markdownBlock) {
                    RenderFallback(markdownBlock);
                }
            }

            protected override void VisitHeadingBlock(Omd.HeadingBlock block) {
                var headingParagraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(headingParagraph, _quoteDepth, _alignment);
                ProcessInlinesOmd(block.Inlines, headingParagraph, _options, _document, _converter._currentFootnotes, _pageContentWidthPixels, _listLevel, _quoteDepth);
                headingParagraph.Style = HeadingStyleMapper.GetHeadingStyleForLevel(block.Level);
                ApplyHeadingTheme(headingParagraph, _options);
            }

            protected override void VisitParagraphBlock(Omd.ParagraphBlock block) {
                var paragraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                ProcessInlinesOmd(block.Inlines, paragraph, _options, _document, _converter._currentFootnotes, _pageContentWidthPixels, _listLevel, _quoteDepth);
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
                    ApplyBodyTextTheme(captionParagraph, _options);
                }
            }

            protected override void VisitCodeBlock(Omd.CodeBlock block) {
                var codeParagraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(codeParagraph, _quoteDepth, _alignment);
                var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";
                codeParagraph.AddFormattedText(block.Content ?? string.Empty).SetFontFamily(monoFont);
                ApplyCodeTheme(codeParagraph, _options);
                if (!string.IsNullOrWhiteSpace(block.Caption)) {
                    var captionParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(captionParagraph, _quoteDepth, _alignment);
                    captionParagraph.AddText(block.Caption!);
                    ApplyBodyTextTheme(captionParagraph, _options);
                }
            }

            protected override void VisitSemanticFencedBlock(Omd.SemanticFencedBlock block) {
                if (TryRenderWordPageBreakSemanticBlock(block, _host, _quoteDepth, _alignment)) {
                    return;
                }

                if (_converter.TryRenderWordHeaderFooterSemanticBlock(block, _host, _options, _document, _pageContentWidthPixels)) {
                    return;
                }

                var paragraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";
                paragraph.AddFormattedText(block.Content ?? string.Empty).SetFontFamily(monoFont);
                ApplyCodeTheme(paragraph, _options);
                if (!string.IsNullOrWhiteSpace(block.Caption)) {
                    var captionParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(captionParagraph, _quoteDepth, _alignment);
                    captionParagraph.AddText(block.Caption!);
                    ApplyBodyTextTheme(captionParagraph, _options);
                }
            }

            protected override void VisitTableBlock(Omd.TableBlock block) =>
                _converter.RenderSharedTableBlockOmd(block, _host, _options, _document, _pageContentWidthPixels);

            protected override void VisitUnorderedListBlock(Omd.UnorderedListBlock block) =>
                RenderListBlock(block.Items, WordListStyle.Bulleted, startNumber: null);

            protected override void VisitOrderedListBlock(Omd.OrderedListBlock block) =>
                RenderListBlock(block.Items, WordListStyle.Numbered, block.Start);

            protected override void VisitTocBlock(Omd.TocBlock block) {
                if (block.Scope != Omd.TocScope.Document) {
                    RenderTocFallback(block);
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

                RenderTocFallback(block);
            }

            private void RenderTocFallback(Omd.TocBlock block) {
                RenderTocFallbackTitle(block);

                if (block.Entries.Count == 0) {
                    return;
                }

                var list = _host.CreateList(block.Ordered ? WordListStyle.Numbered : WordListStyle.Bulleted);
                int baseLevel = block.NormalizeLevels ? block.Entries.Min(entry => entry.Level) : 1;
                foreach (var entry in block.Entries) {
                    if (string.IsNullOrWhiteSpace(entry.Text)) {
                        continue;
                    }

                    int effectiveLevel = Math.Max(0, _listLevel + entry.Level - baseLevel);
                    var paragraph = list.AddItem((string?)null, effectiveLevel);
                    ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                    if (!string.IsNullOrWhiteSpace(entry.Anchor)) {
                        paragraph.AddHyperLink(entry.Text, entry.Anchor.TrimStart('#'), addStyle: true);
                    } else {
                        paragraph.AddText(entry.Text);
                        ApplyBodyTextTheme(paragraph, _options);
                    }
                }

                _host.NotifyListRendered(list);
            }

            private void RenderTocFallbackTitle(Omd.TocBlock block) {
                if (!block.IncludeTitle || block.TitleHeadingAlreadyRendered || string.IsNullOrWhiteSpace(block.Title)) {
                    return;
                }

                var headingParagraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(headingParagraph, _quoteDepth, _alignment);
                headingParagraph.AddText(block.Title.Trim());
                headingParagraph.Style = HeadingStyleMapper.GetHeadingStyleForLevel(NormalizeTocTitleLevel(block.TitleLevel));
                ApplyHeadingTheme(headingParagraph, _options);
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
                } else if (_converter.TryRenderHtmlFallbackViaMarkdownAst(block.Html, _host, _options, _document, _quoteDepth, _pageContentWidthPixels, _alignment)) {
                    return;
                } else {
                    var htmlParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(htmlParagraph, _quoteDepth, _alignment);
                    htmlParagraph.AddText(((Omd.IMarkdownBlock)block).RenderMarkdown());
                    ApplyBodyTextTheme(htmlParagraph, _options);
                }
            }

            protected override void VisitHorizontalRuleBlock(Omd.HorizontalRuleBlock block) {
                if (_host.SupportsHorizontalRule) {
                    _host.InsertHorizontalRule();
                } else {
                    var ruleParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(ruleParagraph, _quoteDepth, _alignment);
                    ruleParagraph.AddText("---");
                    ApplyBodyTextTheme(ruleParagraph, _options);
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

                    _converter.RenderSharedDefinitionListEntryOmd(entry, _host, _options, _document, _quoteDepth, _pageContentWidthPixels, _alignment);
                }
            }

            protected override void VisitQuoteBlock(Omd.QuoteBlock block) {
                foreach (var child in block.Children) {
                    RenderNested(child, quoteDepth: _quoteDepth + 1);
                }
            }

            protected override void VisitCalloutBlock(Omd.CalloutBlock block) =>
                _converter.RenderSharedCalloutBlockOmd(block, _host, _options, _document, _quoteDepth, _pageContentWidthPixels, _alignment);

            protected override void VisitFootnoteDefinitionBlock(Omd.FootnoteDefinitionBlock block) { }

            protected override void VisitDetailsBlock(Omd.DetailsBlock block) {
                if (block.Summary != null) {
                    var summaryParagraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(summaryParagraph, _quoteDepth, _alignment);
                    ProcessInlinesOmd(block.Summary.Inlines, summaryParagraph, _options, _document, _converter._currentFootnotes, _pageContentWidthPixels, _listLevel, _quoteDepth);
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
                ProcessInlinesOmd(block.Inlines, summaryParagraph, _options, _document, _converter._currentFootnotes, _pageContentWidthPixels, _listLevel, _quoteDepth);
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
                    ApplyCodeTheme(paragraph, _options);
                }
            }

            private static int NormalizeTocLevel(int level, int fallback) {
                int normalized = level <= 0 ? fallback : level;
                if (normalized < 1) {
                    return 1;
                }

                return normalized > 9 ? 9 : normalized;
            }

            private static int NormalizeTocTitleLevel(int level) {
                if (level < 1) {
                    return Omd.TocOptions.DefaultTitleLevel;
                }

                return level > 6 ? 6 : level;
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
                            ProcessInlinesOmd(paragraph.Inlines, listItemParagraph, _options, _document, _converter._currentFootnotes, _pageContentWidthPixels, effectiveLevel + 1, _quoteDepth);
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

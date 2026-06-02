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
                ProcessInlinesOmd(block.Inlines, headingParagraph, _options, _document, _currentFootnotes);
                headingParagraph.Style = HeadingStyleMapper.GetHeadingStyleForLevel(block.Level);
            }

            protected override void VisitParagraphBlock(Omd.ParagraphBlock block) {
                var paragraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                ProcessInlinesOmd(block.Inlines, paragraph, _options, _document, _currentFootnotes);
            }

            protected override void VisitImageBlock(Omd.ImageBlock block) {
                var paragraph = _host.CreateParagraph();
                ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                var pathOrUrl = block.Path ?? string.Empty;
                var contextWidthLimit = ResolveContextWidthLimitPixels(_options.ImageLayout, _pageContentWidthPixels, _listLevel, _quoteDepth);

                if (System.IO.File.Exists(pathOrUrl)) {
                    if (_options.AllowLocalImages && LocalPathAllowed(pathOrUrl, _options)) {
                        double? naturalW = null;
                        double? naturalH = null;
                        if (TryGetImageDimensionsFromFile(pathOrUrl, out var fileW, out var fileH)) {
                            naturalW = fileW;
                            naturalH = fileH;
                        }

                        ResolveImageDimensions(
                            _options,
                            source: pathOrUrl,
                            context: "block-local",
                            requestedWidth: block.Width,
                            requestedHeight: block.Height,
                            naturalWidth: naturalW,
                            naturalHeight: naturalH,
                            pageContentWidthPixels: _pageContentWidthPixels,
                            contextWidthLimitPixels: contextWidthLimit,
                            out var finalW,
                            out var finalH,
                            out _);

                        paragraph.AddImage(pathOrUrl, finalW, finalH, description: block.Alt ?? string.Empty);
                    } else {
                        var text = paragraph.AddText(block.Alt ?? System.IO.Path.GetFileName(pathOrUrl));
                        var defaultFont = ResolveDefaultFontFamily(_options);
                        if (!string.IsNullOrEmpty(defaultFont)) {
                            text.SetFontFamily(defaultFont!);
                        }
                    }
                } else if (System.Uri.TryCreate(pathOrUrl, System.UriKind.Absolute, out var uri)) {
                    if (_options.AllowedImageSchemes.Contains(uri.Scheme) &&
                        (_options.ImageUrlValidator == null || _options.ImageUrlValidator(uri))) {
                        if (_options.AllowRemoteImages) {
                            try {
                                var bytes = DownloadRemoteImageBytes(uri, _options);
                                var fileName = System.IO.Path.GetFileName(uri.LocalPath);
                                if (string.IsNullOrWhiteSpace(fileName)) {
                                    fileName = "image";
                                }

                                double? naturalW = null;
                                double? naturalH = null;
                                if (TryGetImageDimensionsFromBytes(bytes, out var remoteW, out var remoteH)) {
                                    naturalW = remoteW;
                                    naturalH = remoteH;
                                }

                                ResolveImageDimensions(
                                    _options,
                                    source: uri.ToString(),
                                    context: "block-remote",
                                    requestedWidth: block.Width,
                                    requestedHeight: block.Height,
                                    naturalWidth: naturalW,
                                    naturalHeight: naturalH,
                                    pageContentWidthPixels: _pageContentWidthPixels,
                                    contextWidthLimitPixels: contextWidthLimit,
                                    out var finalW,
                                    out var finalH,
                                    out _);

                                using var stream = new System.IO.MemoryStream(bytes, writable: false);
                                paragraph.AddImage(stream, fileName, finalW, finalH, description: block.Alt ?? string.Empty);
                            } catch (Exception ex) {
                                _options.OnWarning?.Invoke($"Remote image '{uri}' could not be downloaded. {ex.Message}");
                                if (_options.FallbackRemoteImagesToHyperlinks) {
                                    paragraph.AddHyperLink(block.Alt ?? uri.ToString(), uri);
                                }
                            }
                        } else if (_options.FallbackRemoteImagesToHyperlinks) {
                            paragraph.AddHyperLink(block.Alt ?? uri.ToString(), uri);
                        }
                    } else if (_options.FallbackRemoteImagesToHyperlinks) {
                        paragraph.AddHyperLink(block.Alt ?? uri.ToString(), uri);
                    }
                } else {
                    var text = paragraph.AddText(block.Alt ?? pathOrUrl);
                    var defaultFont = ResolveDefaultFontFamily(_options);
                    if (!string.IsNullOrEmpty(defaultFont)) {
                        text.SetFontFamily(defaultFont!);
                    }
                }

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

            protected override void VisitTocBlock(Omd.TocBlock block) { }

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
                    ProcessInlinesOmd(block.Summary.Inlines, summaryParagraph, _options, _document, _currentFootnotes);
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
                ProcessInlinesOmd(block.Inlines, summaryParagraph, _options, _document, _currentFootnotes);
                foreach (var run in summaryParagraph.GetRuns()) {
                    run.SetBold();
                }
            }

            protected override void VisitFrontMatterBlock(Omd.FrontMatterBlock block) {
                var lines = block.Render().Replace("\r", string.Empty).Split('\n');
                var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";

                for (int i = 0; i < lines.Length; i++) {
                    var paragraph = _host.CreateParagraph();
                    ApplyBlockParagraphFormatting(paragraph, _quoteDepth, _alignment);
                    paragraph.AddFormattedText(lines[i]).SetFontFamily(monoFont);
                }
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
                            ProcessInlinesOmd(paragraph.Inlines, listItemParagraph, _options, _document, _currentFootnotes);
                            firstParagraph = false;
                            continue;
                        }

                        RenderNested(blockChildren[i], listLevel: effectiveLevel + 1);
                    }
                }
            }
        }
    }
}

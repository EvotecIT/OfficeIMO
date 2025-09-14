using Omd = OfficeIMO.Markdown;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Markdown → Word converter powered by OfficeIMO.Markdown.
    /// Maps OMD blocks/inlines onto OfficeIMO.Word APIs (headings, lists, tables, images,
    /// code, quotes, callouts, footnotes, etc.).
    /// </summary>
    internal partial class MarkdownToWordConverter {
        public WordDocument Convert(string markdown, MarkdownToWordOptions options) {
            return ConvertAsync(markdown, options).GetAwaiter().GetResult();
        }

        public Task<WordDocument> ConvertAsync(string markdown, MarkdownToWordOptions options, CancellationToken cancellationToken = default) {
            if (markdown == null) {
                throw new ArgumentNullException(nameof(markdown));
            }

            options ??= new MarkdownToWordOptions();

            var document = WordDocument.Create();
            options.ApplyDefaults(document);

            // Parse using OfficeIMO.Markdown reader.
            var omd = Omd.MarkdownReader.Parse(markdown);
            // Build footnote definitions map for this document
            _currentFootnotes = omd.Blocks is not null
                ? omd.Blocks
                    .OfType<Omd.FootnoteDefinitionBlock>()
                    .GroupBy(f => f.Label)
                    .ToDictionary(g => g.Key, g => g.Last().Text)
                : null;
            foreach (var block in omd.Blocks) {
                cancellationToken.ThrowIfCancellationRequested();
                ProcessBlockOmd(block, document, options, quoteDepth: 0);
            }

            return Task.FromResult(document);
        }

        //

        // New OfficeIMO.Markdown path
        private static void ProcessBlockOmd(Omd.IMarkdownBlock block, WordDocument document, MarkdownToWordOptions options, WordList? currentList = null, int listLevel = 0, int quoteDepth = 0) {
            switch (block) {
                case Omd.HeadingBlock h:
                    var headingParagraph = document.AddParagraph(h.Text ?? string.Empty);
                    if (quoteDepth > 0) headingParagraph.IndentationBefore = 720 * quoteDepth;
                    headingParagraph.Style = HeadingStyleMapper.GetHeadingStyleForLevel(h.Level);
                    break;
                case Omd.ParagraphBlock p:
                    if (currentList == null) {
                        var para = document.AddParagraph(string.Empty);
                        if (quoteDepth > 0) para.IndentationBefore = 720 * quoteDepth;
                        ProcessInlinesOmd(p.Inlines, para, options, document, _currentFootnotes);
                    } else {
                        var li = currentList.AddItem(string.Empty, listLevel);
                        if (quoteDepth > 0) li.IndentationBefore = 720 * quoteDepth;
                        ProcessInlinesOmd(p.Inlines, li, options, document, _currentFootnotes);
                    }
                    break;
                case Omd.ImageBlock img:
                    {
                        var par = document.AddParagraph(string.Empty);
                        var pathOrUrl = img.Path ?? string.Empty;
                        double? w = img.Width;
                        double? h = img.Height;
                        if (System.IO.File.Exists(pathOrUrl)) {
                            if (w == null || h == null) {
                                try {
                                    using var image = SixLabors.ImageSharp.Image.Load(pathOrUrl, out _);
                                    w ??= image.Width;
                                    h ??= image.Height;
                                } catch { }
                            }
                            par.AddImage(pathOrUrl, w, h, description: img.Alt ?? string.Empty);
                        } else if (pathOrUrl.StartsWith("http", StringComparison.OrdinalIgnoreCase)) {
                            document.AddImageFromUrl(pathOrUrl, w, h).Description = img.Alt ?? string.Empty;
                        } else {
                            var hl = par.AddHyperLink(img.Alt ?? pathOrUrl, new System.Uri(pathOrUrl, System.UriKind.RelativeOrAbsolute));
                            if (!string.IsNullOrEmpty(options.FontFamily)) hl.SetFontFamily(options.FontFamily!);
                        }
                        if (!string.IsNullOrWhiteSpace(img.Caption)) document.AddParagraph(img.Caption!);
                        break;
                    }
                case Omd.CodeBlock cb:
                    var codeParagraph = document.AddParagraph(string.Empty);
                    if (quoteDepth > 0) codeParagraph.IndentationBefore = 720 * quoteDepth;
                    var monoFont = FontResolver.Resolve("monospace") ?? "Consolas";
                    codeParagraph.AddFormattedText(cb.Content ?? string.Empty).SetFontFamily(monoFont);
                    if (!string.IsNullOrWhiteSpace(cb.Caption)) document.AddParagraph(cb.Caption!);
                    break;
                case Omd.TableBlock tb:
                    // Create table and map cell alignments + inline formatting inside cells
                    var cols = (tb.Headers.Count > 0 ? tb.Headers.Count : (tb.Rows.Count > 0 ? tb.Rows[0].Count : 1));
                    var rows = tb.Rows.Count + (tb.Headers.Count > 0 ? 1 : 0);
                    var wtable = document.AddTable(rows: rows, columns: cols);
                    int r = 0;
                    if (tb.Headers.Count > 0) {
                        for (int c = 0; c < tb.Headers.Count; c++) {
                            var para = wtable.Rows[r].Cells[c].Paragraphs[0];
                            ProcessInlinesOmd(Omd.MarkdownReader.ParseInlineText(tb.Headers[c]), para, options, document, _currentFootnotes);
                            // Apply header alignment if provided
                            if (c < tb.Alignments.Count) ApplyAlignment(tb.Alignments[c], para);
                        }
                        r++;
                    }
                    foreach (var row in tb.Rows) {
                        for (int c = 0; c < row.Count && c < wtable.Rows[r].Cells.Count; c++) {
                            var para = wtable.Rows[r].Cells[c].Paragraphs[0];
                            ProcessInlinesOmd(Omd.MarkdownReader.ParseInlineText(row[c]), para, options, document, _currentFootnotes);
                            if (c < tb.Alignments.Count) ApplyAlignment(tb.Alignments[c], para);
                        }
                        r++;
                    }
                    break;
                case Omd.UnorderedListBlock ul:
                    {
                        var list = document.AddList(WordListStyle.Bulleted);
                        foreach (var item in ul.Items) {
                            var li = list.AddItem(string.Empty, listLevel);
                            // Task list support
                            if (item.IsTask) li.AddCheckBox(item.Checked);
                            ProcessInlinesOmd(item.Content, li, options, document, _currentFootnotes);
                        }
                        break;
                    }
                case Omd.OrderedListBlock ol:
                    {
                        var list = document.AddList(WordListStyle.Numbered);
                        if (ol.Start != 1) list.Numbering.Levels[0].SetStartNumberingValue(ol.Start);
                        foreach (var item in ol.Items) {
                            var li = list.AddItem(string.Empty, listLevel);
                            ProcessInlinesOmd(item.Content, li, options, document, _currentFootnotes);
                        }
                        break;
                    }
                case Omd.TocBlock:
                    // Skip TOC for Word
                    break;
                case Omd.HtmlRawBlock html:
                    document.AddHtmlToBody(html.Html);
                    break;
                case Omd.HorizontalRuleBlock:
                    document.AddHorizontalLine();
                    break;
                case Omd.QuoteBlock qb:
                    foreach (var child in qb.Children) ProcessBlockOmd(child, document, options, null, 0, quoteDepth + 1);
                    break;
                case Omd.CalloutBlock callout:
                    // Render as a simple bold title followed by body paragraphs
                    var ptitle = document.AddParagraph(string.Empty);
                    ptitle.AddFormattedText(callout.Title, bold: true);
                    var pbody = document.AddParagraph(callout.Body);
                    break;
                case Omd.FootnoteDefinitionBlock:
                    // Definitions are consumed when encountering references; skip emitting as body paragraphs.
                    break;
                default:
                    // Fallback: render markdown text
                    var txt = ((Omd.IMarkdownBlock)block).RenderMarkdown();
                    document.AddParagraph(txt);
                    break;
            }
        }

        // Current footnote definitions map; built per-document on ConvertAsync
        private static IReadOnlyDictionary<string, string>? _currentFootnotes;

        private static void ApplyAlignment(Omd.ColumnAlignment align, WordParagraph para) {
            switch (align) {
                case Omd.ColumnAlignment.Left: para.ParagraphAlignment = JustificationValues.Left; break;
                case Omd.ColumnAlignment.Center: para.ParagraphAlignment = JustificationValues.Center; break;
                case Omd.ColumnAlignment.Right: para.ParagraphAlignment = JustificationValues.Right; break;
            }
        }
    }
}

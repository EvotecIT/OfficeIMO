using System.Linq;
using OfficeIMO.Markdown;
using MarkdigMarkdown = Markdig.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_Quote_Content_Tests {
        [Fact]
        public void Quote_With_List_And_TaskItems() {
            var md = string.Join("\n", new[] {
                "> Shopping",
                ">",
                "> - Apples",
                "> - Bananas",
                "> - [x] Done"
            });
            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            Assert.Same(qb.Children, qb.ChildBlocks);
            Assert.True(qb.Children.Count >= 2);
            Assert.IsType<ParagraphBlock>(qb.Children[0]);
            var ul = Assert.IsType<UnorderedListBlock>(qb.Children[1]);
            Assert.Equal(3, ul.Items.Count);
            Assert.True(ul.Items[2].IsTask);
            Assert.True(ul.Items[2].Checked);
            var round = doc.ToMarkdown();
            Assert.Contains("> - [x] Done", round);
        }

        [Fact]
        public void Quote_With_CodeBlock_And_Caption() {
            var md = string.Join("\n", new[] {
                "> ```csharp",
                "> Console.WriteLine(\"Hi\");",
                "> ```",
                "> _Sample_"
            });
            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            var code = Assert.IsType<CodeBlock>(qb.Children[0]);
            Assert.Equal("csharp", code.Language);
            Assert.Equal("Sample", code.Caption);
            var outMd = doc.ToMarkdown();
            Assert.Contains("> ```csharp", outMd);
            Assert.Contains("> _Sample_", outMd);
        }

        [Fact]
        public void Quote_Nested_Quote_With_List() {
            var md = string.Join("\n", new[] {
                "> Outer",
                "> > Inner",
                "> > - a",
                "> > - b",
                "> After"
            });
            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            Assert.Equal(2, qb.Children.Count);
            var inner = Assert.IsType<QuoteBlock>(qb.Children[1]);
            Assert.Equal(2, inner.Children.Count);
            var list = Assert.IsType<UnorderedListBlock>(inner.Children[1]);
            Assert.Equal(2, list.Items.Count);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<blockquote><p>Outer</p><blockquote><p>Inner</p><ul><li>a</li><li>b\nAfter</li></ul></blockquote></blockquote>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Quote_With_Table() {
            var md = string.Join("\n", new[] {
                "> | A | B |",
                "> | --- | ---: |",
                "> | 1 | 2 |"
            });
            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            var table = Assert.IsType<TableBlock>(qb.ChildBlocks.First());
            Assert.Equal(new[] { "A", "B" }, table.Headers);
            Assert.Equal(new[] { ColumnAlignment.None, ColumnAlignment.Right }, table.Alignments);
        }

        [Fact]
        public void Quote_Lazy_Continuation_Extends_Unordered_List_Item() {
            const string md = "> - item\ncontinuation";

            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            var list = Assert.IsType<UnorderedListBlock>(qb.Children.First());
            Assert.Single(list.Items);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<blockquote><ul><li>item continuation</li></ul></blockquote>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Quote_Lazy_Paragraph_Continuation_Preserves_Markdig_SoftBreak() {
            const string md = "> quote\nlazy\n";

            var result = MarkdownReader.ParseWithSyntaxTree(md);
            var quote = Assert.IsType<QuoteBlock>(Assert.Single(result.Document.Blocks));
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));
            var quoteSyntax = Assert.Single(result.SyntaxTree.Children);
            var paragraphSyntax = Assert.Single(quoteSyntax.Children, child => child.Kind == MarkdownSyntaxKind.Paragraph);
            var written = NormalizeMarkdown(result.Document.ToMarkdown());
            var office = result.Document.ToHtmlFragment(CreatePlainHtmlOptions());
            var reparsedOffice = MarkdownReader.Parse(written).ToHtmlFragment(CreatePlainHtmlOptions());
            var markdig = MarkdigMarkdown.ToHtml(md);

            Assert.Equal("quote\nlazy", paragraph.Inlines.RenderMarkdown());
            Assert.IsType<SoftBreakInline>(paragraph.Inlines.Nodes[1]);
            Assert.Equal(
                new[] {
                    MarkdownSyntaxKind.InlineText,
                    MarkdownSyntaxKind.InlineSoftBreak,
                    MarkdownSyntaxKind.InlineText
                },
                paragraphSyntax.Children.Select(child => child.Kind).ToArray());
            Assert.Equal(new MarkdownSourceSpan(1, 3, 2, 4), paragraph.SourceSpan);
            Assert.Equal(new MarkdownSourceSpan(1, 3, 2, 4), paragraphSyntax.SourceSpan);
            Assert.Equal("> quote\n> lazy", written);
            Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
            Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));

            var native = MarkdownNativeDocument.Parse(md);
            var nativeQuote = Assert.IsType<MarkdownNativeQuoteBlock>(Assert.Single(native.Blocks));
            var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(nativeQuote.Children));
            Assert.Equal(new MarkdownSourceSpan(1, 3, 2, 4), nativeQuote.BodySourceSpan);
            Assert.Contains(nativeParagraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.SoftBreak);
            MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        }

        [Fact]
        public void Quote_Lazy_Paragraph_Setext_Looking_Line_Stays_Paragraph_Text() {
            const string md = "> foo\nbar\n===\n";

            var result = MarkdownReader.ParseWithSyntaxTree(md, MarkdownReaderOptions.CreateCommonMarkProfile());
            var quote = Assert.IsType<QuoteBlock>(Assert.Single(result.Document.Blocks));
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));
            var quoteSyntax = Assert.Single(result.FinalSyntaxTree.Children);
            var paragraphSyntax = Assert.Single(quoteSyntax.Children, child => child.Kind == MarkdownSyntaxKind.Paragraph);
            var written = NormalizeMarkdown(result.Document.ToMarkdown());
            var office = result.Document.ToHtmlFragment(CreatePlainHtmlOptions());
            var reparsedOffice = MarkdownReader.Parse(written, MarkdownReaderOptions.CreateCommonMarkProfile()).ToHtmlFragment(CreatePlainHtmlOptions());
            var markdig = MarkdigMarkdown.ToHtml(md);

            Assert.Equal("foo\nbar\n===", paragraph.Inlines.RenderMarkdown().Replace("\r\n", "\n"));
            Assert.DoesNotContain(quote.ChildBlocks, child => child is HeadingBlock);
            Assert.DoesNotContain(quoteSyntax.Descendants(), node => node.Kind == MarkdownSyntaxKind.Heading);
            Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraphSyntax.Kind);
            Assert.Equal("> foo\n> bar\n> \\===", written);
            Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(office));
            Assert.Equal(NormalizeHtml(markdig), NormalizeHtml(reparsedOffice));
            MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        }

        [Fact]
        public void Quote_Lazy_Continuation_Extends_Ordered_List_Item() {
            const string md = "> 1. item\ncontinuation";

            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            var list = Assert.IsType<OrderedListBlock>(qb.Children.First());
            Assert.Single(list.Items);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<blockquote><ol><li>item continuation</li></ol></blockquote>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Quote_Explicit_Continuation_Extends_Ordered_List_Item() {
            const string md = "> 1. item\n>   continuation";

            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            var list = Assert.IsType<OrderedListBlock>(qb.Children.First());
            Assert.Single(list.Items);

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<blockquote><ol><li>item continuation</li></ol></blockquote>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Quote_Indented_Paragraph_Line_Can_Stay_In_Paragraph_And_Allow_Lazy_Continuation() {
            const string md = "> quote\n>     code\ncontinuation";

            var doc = MarkdownReader.Parse(md);
            var qb = Assert.IsType<QuoteBlock>(doc.Blocks[0]);
            var paragraph = Assert.IsType<ParagraphBlock>(qb.Children.First());
            Assert.Equal("quote code\ncontinuation", paragraph.Inlines.RenderMarkdown());

            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<blockquote><p>quote code\ncontinuation</p></blockquote>", html, StringComparison.Ordinal);
        }

        private static HtmlOptions CreatePlainHtmlOptions() => new() {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            AutoHeadingIdentifiers = false
        };

        private static string NormalizeHtml(string html) {
            if (string.IsNullOrWhiteSpace(html)) {
                return string.Empty;
            }

            var compact = html
                .Replace("\r\n", "\n")
                .Replace('\r', '\n')
                .Replace("> <", "><")
                .Trim();
            var sb = new System.Text.StringBuilder(compact.Length);
            bool lastWasWhitespace = false;
            for (int i = 0; i < compact.Length; i++) {
                char ch = compact[i];
                if (char.IsWhiteSpace(ch)) {
                    lastWasWhitespace = true;
                    continue;
                }

                if (lastWasWhitespace && sb.Length > 0 && sb[sb.Length - 1] != '>') {
                    sb.Append(' ');
                }

                lastWasWhitespace = false;
                sb.Append(ch);
            }

            return sb.ToString();
        }

        private static string NormalizeMarkdown(string markdown) =>
            markdown
                .Replace("\r\n", "\n")
                .Replace('\r', '\n')
                .TrimEnd('\n');
    }
}

using OfficeIMO.Latex;
using OfficeIMO.Latex.Markdown;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexStructuredContentTests {
    private static MarkdownDoc Convert(string body) => LatexDocument.Parse(
        "\\documentclass{article}\\begin{document}" + body + "\\end{document}").ToMarkdownDocumentResult().Value;

    [Fact]
    public void List_items_own_nested_quotations_and_following_paragraphs_in_source_order() {
        MarkdownDoc document = Convert("\\begin{itemize}\\item First\n\n\\begin{quote}Quoted\\end{quote}\n\nContinued\\item Last\\end{itemize}");
        UnorderedListBlock list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        Assert.Equal(2, list.Items.Count);
        Assert.Collection(list.Items[0].ChildBlocks,
            block => Assert.Equal("First", Plain(Assert.IsType<ParagraphBlock>(block).Inlines)),
            block => Assert.IsType<QuoteBlock>(block),
            block => Assert.Equal("Continued", Plain(Assert.IsType<ParagraphBlock>(block).Inlines)));
        Assert.Equal("Last", Plain(list.Items[1].Content));
        Assert.Contains("Continued", document.ToMarkdown());
    }

    [Fact]
    public void Quotations_own_nested_lists_without_promoting_items_to_outer_list() {
        MarkdownDoc document = Convert(@"\begin{quotation}Intro\begin{enumerate}\item One\begin{itemize}\item Nested\end{itemize}\item Two\end{enumerate}Tail\end{quotation}");
        QuoteBlock quote = Assert.IsType<QuoteBlock>(Assert.Single(document.Blocks));
        Assert.Equal(3, quote.ChildBlocks.Count);
        OrderedListBlock ordered = Assert.IsType<OrderedListBlock>(quote.ChildBlocks[1]);
        Assert.Equal(2, ordered.Items.Count);
        UnorderedListBlock nested = Assert.IsType<UnorderedListBlock>(Assert.Single(ordered.Items[0].NestedBlocks));
        Assert.Equal("Nested", Plain(Assert.Single(nested.Items).Content));
        Assert.Equal("Two", Plain(ordered.Items[1].Content));
        Assert.Contains("Tail", document.ToMarkdown());
    }

    [Fact]
    public void Description_definitions_keep_structured_children_and_multiple_paragraphs() {
        MarkdownDoc document = Convert("\\begin{description}\\item[Term]First\n\nSecond\\begin{quote}Quoted\\end{quote}\\end{description}");
        DefinitionListBlock definitions = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));
        DefinitionListEntry entry = Assert.Single(definitions.Entries);
        Assert.Equal("Term", Plain(entry.Term));
        Assert.Collection(entry.DefinitionBlocks,
            block => Assert.Equal("First", Plain(Assert.IsType<ParagraphBlock>(block).Inlines)),
            block => Assert.Equal("Second", Plain(Assert.IsType<ParagraphBlock>(block).Inlines)),
            block => Assert.IsType<QuoteBlock>(block));
    }

    [Fact]
    public void Commands_wrapping_nested_blocks_stay_complete_source_in_container_content() {
        const string wrapped = @"\resizebox{1}{!}{\begin{itemize}\item Inert\end{itemize}}";
        MarkdownDoc document = Convert(@"\begin{quote}Before " + wrapped + @" After\end{quote}");
        QuoteBlock quote = Assert.IsType<QuoteBlock>(Assert.Single(document.Blocks));
        CodeBlock fallback = Assert.Single(quote.ChildBlocks.OfType<CodeBlock>());
        Assert.Equal(wrapped, fallback.Content);
        Assert.DoesNotContain(quote.ChildBlocks, block => block is UnorderedListBlock);
        Assert.Contains("Before", document.ToMarkdown());
        Assert.Contains("After", document.ToMarkdown());
    }

    [Fact]
    public void Required_single_token_arguments_convert_with_following_text_outside_the_formatting() {
        MarkdownDoc document = Convert(@"\textbf XYZ and \emph\% tail.");
        ParagraphBlock paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        Assert.Equal("XYZ and % tail.", Plain(paragraph.Inlines));
        Assert.Equal("X", Plain(Assert.Single(paragraph.Inlines.Nodes.OfType<BoldSequenceInline>())));
        Assert.Contains("**X**YZ", document.ToMarkdown());
    }

    private static string Plain(IPlainTextMarkdownInline value) {
        var output = new System.Text.StringBuilder();
        value.AppendPlainText(output);
        return output.ToString();
    }
}

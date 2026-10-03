using OfficeIMO.Latex;
using OfficeIMO.Latex.Markdown;
using OfficeIMO.Latex.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Latex.Markdown.Tests;

public class LatexFootnoteConversionTests {
    [Fact]
    public void RichNoteBodyDoesNotFragmentItsSurroundingParagraph() {
        const string body = "Before\\footnote{First \\emph{note}.\n\nSecond.\\begin{itemize}\\item One\\item Two\\end{itemize}\\begin{quote}Quoted\\end{quote}} after.";
        LatexDocument source = LatexDocument.Parse(Wrap(body));
        Assert.Equal(body, Assert.Single(source.Paragraphs).Content);
        LatexToMarkdownResult result = source.ToMarkdownDocumentResult();
        ParagraphBlock paragraph = Assert.IsType<ParagraphBlock>(result.Value.Blocks[0]);
        FootnoteRefInline reference = Assert.Single(paragraph.Inlines.Nodes.OfType<FootnoteRefInline>());
        FootnoteDefinitionBlock definition = Assert.IsType<FootnoteDefinitionBlock>(result.Value.Blocks[1]);
        Assert.Equal(reference.Label, definition.Label);
        Assert.Collection(definition.ChildBlocks,
            block => Assert.Contains("First", Plain(Assert.IsType<ParagraphBlock>(block).Inlines)),
            block => Assert.Equal("Second.", Plain(Assert.IsType<ParagraphBlock>(block).Inlines)),
            block => Assert.Equal(2, Assert.IsType<UnorderedListBlock>(block).Items.Count),
            block => Assert.IsType<QuoteBlock>(block));
        Assert.Single(Assert.IsType<ParagraphBlock>(definition.ChildBlocks[0]).Inlines.Nodes.OfType<ItalicSequenceInline>());
        Assert.Contains(" after.", Plain(paragraph.Inlines), StringComparison.Ordinal);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Feature.StartsWith("command:footnote", StringComparison.Ordinal));
        string pdfText = PdfReadDocument.Open(source.ToPdfDocumentResult().Value.ToBytes()).ExtractText();
        foreach (string text in new[] { "Before", "after.", "First", "Second.", "One", "Two", "Quoted" })
            Assert.Single(System.Text.RegularExpressions.Regex.Matches(pdfText,
                System.Text.RegularExpressions.Regex.Escape(text)).Cast<System.Text.RegularExpressions.Match>());
    }

    [Fact]
    public void NotesInsideStructuredContainersShareUniqueDefinitionsAndKeepExplicitMarkLossVisible() {
        LatexToMarkdownResult result = LatexDocument.Parse(Wrap("\\begin{itemize}\\item First\\footnote[9]{Alpha}\\begin{quote}Second\\footnote B tail\\end{quote}\\end{itemize}")).ToMarkdownDocumentResult();
        FootnoteDefinitionBlock[] definitions = result.Value.Blocks.OfType<FootnoteDefinitionBlock>().ToArray();
        Assert.Equal(2, definitions.Length);
        Assert.NotEqual(definitions[0].Label, definitions[1].Label);
        Assert.Equal(new[] { "Alpha", "B" }, definitions.Select(definition => Plain(Assert.Single(definition.ChildBlocks.OfType<ParagraphBlock>()).Inlines)));
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Feature == "footnote-mark" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Simplified);
        Assert.Equal(2, result.Value.Descendants().OfType<FootnoteRefInline>().Count());
        Assert.DoesNotContain(result.Value.ToMarkdown(), "\\footnote", StringComparison.Ordinal);
    }

    [Fact]
    public void FootnotesInOpaqueArgumentsStayInTheFallbackAndNestedInsertionsAreReported() {
        LatexToMarkdownResult result = LatexDocument.Parse(Wrap("\\custom{Hidden\\footnote{Opaque}} Visible\\footnote{Outer\\footnote{Inner}} tail")).ToMarkdownDocumentResult();
        FootnoteDefinitionBlock definition = Assert.Single(result.Value.Blocks.OfType<FootnoteDefinitionBlock>());
        string noteMarkdown = MarkdownDoc.Create().Add(definition).ToMarkdown();
        Assert.Contains("Outer", noteMarkdown, StringComparison.Ordinal);
        Assert.Contains("Inner", noteMarkdown, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Feature == "nested-footnote" && diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback);
        Assert.Contains("Opaque", result.Value.ToMarkdown(), StringComparison.Ordinal);
        Assert.Single(result.Value.Descendants().OfType<FootnoteRefInline>());
    }

    private static string Wrap(string body) => "\\begin{document}\n" + body + "\n\\end{document}";
    private static string Plain(InlineSequence sequence) {
        var text = new System.Text.StringBuilder();
        foreach (IMarkdownInline inline in sequence.Nodes)
            if (inline is IPlainTextMarkdownInline plain) plain.AppendPlainText(text);
        return text.ToString();
    }
}

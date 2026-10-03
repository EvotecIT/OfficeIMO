using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexProjectionContractTests {
    private static string Wrap(string body) => "\\documentclass{article}\n\\begin{document}\n" + body + "\n\\end{document}";

    [Fact]
    public void DisplayScalarsDecodeEscapesAndReportFlattenedFormatting() {
        string body = "\\begin{figure}\\includegraphics{~/a.png}\\caption{\\textbf{R\\&D} 100\\%}\\end{figure}\n" +
            "\\begin{table}\\caption{\\emph{Cost} 50\\%}\\begin{tabular}{l}A\\\\\\end{tabular}\\end{table}\n" +
            "\\begin{theorem}[\\textbf{Useful}~100\\%]Body\\end{theorem}";
        LatexToMarkdownResult result = LatexDocument.Parse("\\title{Cost 100\\%}\\author{\\textbf{Alice\\_Doe}}\\date{R\\&D}" + Wrap(body)).ToMarkdownDocumentResult();
        Assert.Equal("Cost 100%", result.Value.FindFrontMatterEntry("title")?.Value);
        Assert.Equal("Alice_Doe", result.Value.FindFrontMatterEntry("author")?.Value);
        Assert.Equal("R&D", result.Value.FindFrontMatterEntry("date")?.Value);
        Assert.Equal("R&D 100%", Assert.Single(result.Value.Blocks.OfType<ImageBlock>()).Caption);
        Assert.Equal("Cost 50%", Assert.Single(result.Value.Blocks.OfType<TableBlock>()).Attributes.GetAttribute("caption"));
        Assert.Equal("Useful 100%", Assert.Single(result.Value.Blocks.OfType<CalloutBlock>()).Title);
        foreach (string feature in new[] { "metadata:author", "figure-caption", "table-caption", "theorem-title" })
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD115" && diagnostic.Feature == feature);
    }

    [Fact]
    public void PlainMetadataAndFigureCaptionsRoundTripThroughTheWriter() {
        var source = MarkdownDoc.Create().FrontMatter(new Dictionary<string, object?> { ["author"] = "Alice_Doe", ["date"] = "R&D" })
            .Add(new ImageBlock("a.png", "100% useful") { Caption = "100% useful" });
        MarkdownDoc result = source.ToLatexDocument().ToMarkdownDocument();
        Assert.Equal("Alice_Doe", result.FindFrontMatterEntry("author")?.Value);
        Assert.Equal("R&D", result.FindFrontMatterEntry("date")?.Value);
        Assert.Equal("100% useful", Assert.Single(result.Blocks.OfType<ImageBlock>()).Caption);
    }

    [Theory]
    [InlineData("\\textbf{A\\label{x}B}", "**AB**")]
    [InlineData("\\emph{A\\textbf{B\\label{x}C}D}", "**BC**")]
    [InlineData("$A\\label{x}B$", "`AB`")]
    [InlineData("\\unknown{A\\label{x}B}", "\\unknown{AB}")]
    public void TheoremLabelsDoNotSliceEnclosingInlineSyntax(string body, string expected) {
        LatexToMarkdownResult result = LatexDocument.Parse(Wrap("\\begin{theorem}" + body + "\\end{theorem}")).ToMarkdownDocumentResult();
        CalloutBlock callout = Assert.Single(result.Value.Blocks.OfType<CalloutBlock>());
        Assert.Equal("x", callout.Attributes.ElementId);
        Assert.Contains(expected, result.Value.ToMarkdown(), StringComparison.Ordinal);
        Assert.DoesNotContain("\\label", result.Value.ToMarkdown(), StringComparison.Ordinal);
        if (body.StartsWith("\\textbf", StringComparison.Ordinal)) Assert.Empty(result.Report.Diagnostics);
    }

    [Fact]
    public void BatchedBlocksAndFigureImagesHaveCompleteObjectTreeNavigation() {
        MarkdownDoc result = LatexDocument.Parse(Wrap("First\n\n\\begin{figure}\\includegraphics{a.png}\\includegraphics{b.png}\\end{figure}\nLast")).ToMarkdownDocument();
        Assert.Equal(4, result.Blocks.Count);
        for (int index = 0; index < result.Blocks.Count; index++) {
            MarkdownObject block = (MarkdownObject)result.Blocks[index];
            Assert.Same(result, block.Parent);
            Assert.Equal(index + 1, block.IndexInParent); // front matter occupies the first object slot
            Assert.Same(index == 0 ? result.DocumentHeader : result.Blocks[index - 1], block.PreviousSibling);
            Assert.Same(index + 1 == result.Blocks.Count ? null : result.Blocks[index + 1], block.NextSibling);
            foreach (MarkdownObject child in block.Descendants()) Assert.Same(result, child.Document);
        }
    }

    [Theory]
    [InlineData("~")]
    [InlineData("\\textasciitilde{}")]
    public void UrlAndHrefDestinationsRetainLiteralTildes(string tilde) {
        string destination = "https://example.test/" + tilde + "alice";
        MarkdownDoc markdown = LatexDocument.Parse(Wrap("\\url{" + destination + "} \\href{" + destination + "}{Alice}")).ToMarkdownDocument();
        LinkInline[] links = Assert.Single(markdown.Blocks.OfType<ParagraphBlock>()).Inlines.Nodes.OfType<LinkInline>().ToArray();
        Assert.Equal(2, links.Length);
        Assert.All(links, static link => Assert.Equal("https://example.test/~alice", link.Url));
        Assert.Contains("https://example.test/~alice", markdown.ToMarkdown(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InlineAndFigureResourcesRetainTildePaths(bool figure) {
        string image = "\\includegraphics{~/images/a.png}";
        MarkdownDoc markdown = LatexDocument.Parse(Wrap(figure ? "\\begin{figure}" + image + "\\end{figure}" : image)).ToMarkdownDocument();
        string path = figure ? Assert.Single(markdown.Blocks.OfType<ImageBlock>()).Path
            : Assert.Single(Assert.Single(markdown.Blocks.OfType<ParagraphBlock>()).Inlines.Nodes.OfType<ImageInline>()).Src;
        Assert.Equal("~/images/a.png", path);
    }

    [Fact]
    public void ProseStillMapsTildeToSpacingAndExplicitTildeToLiteralText() {
        MarkdownDoc markdown = LatexDocument.Parse(Wrap("Hello~world \\textasciitilde{}")).ToMarkdownDocument();
        var text = new StringBuilder();
        foreach (IPlainTextMarkdownInline inline in Assert.Single(markdown.Blocks.OfType<ParagraphBlock>()).Inlines.Nodes.OfType<IPlainTextMarkdownInline>()) inline.AppendPlainText(text);
        Assert.Equal("Hello world ~", text.ToString());
    }

    [Fact]
    public void IndexedProjectionKeepsCommentsLocalAndOpaquePercentVisible() {
        string body = "\\begin{unknown}First% PRIVATE1\n\\begin{comment}PRIVATE2\\end{comment}Tail\\end{unknown}\n" +
            "Next \\textbf{bold} \\verb|100%|% PRIVATE3\n\n" +
            "\\begin{unknown}Last% PRIVATE4\nEnd\\end{unknown}";
        LatexToMarkdownResult result = LatexDocument.Parse(Wrap(body)).ToMarkdownDocumentResult();
        Assert.Equal(3, result.Value.Blocks.Count);
        string markdown = result.Value.ToMarkdown();
        Assert.DoesNotContain("PRIVATE", markdown, StringComparison.Ordinal);
        Assert.Contains("First", markdown, StringComparison.Ordinal);
        Assert.Contains("Tail", markdown, StringComparison.Ordinal);
        Assert.Contains("Last", markdown, StringComparison.Ordinal);
        Assert.Contains("100%", markdown, StringComparison.Ordinal);
        Assert.Single(result.Report.Diagnostics, static diagnostic => diagnostic.Code == "LATEXMD210");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RangeEnumerationHonorsCancellationAfterItHasStarted(bool comments) {
        LatexDocument document = LatexDocument.Parse(Wrap("\\textbf{a}% one\n\\textit{b}% two\n"));
        using var cancellation = new CancellationTokenSource();
        var context = new LatexProjectionContext(document, cancellation.Token);
        IEnumerable<object> values = comments ? context.Comments(document.SyntaxTree.Root.Span).Cast<object>()
            : context.InlineCandidates(document.Body!.ContentSpan.Start.Offset, document.Body.ContentSpan.End.Offset).Cast<object>();
        using IEnumerator<object> iterator = values.GetEnumerator();
        Assert.True(iterator.MoveNext());
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => iterator.MoveNext());
    }

    [Fact]
    public void SourceExtractionHonorsCancellationDuringInputEnumeration() {
        LatexDocument document = LatexDocument.Parse(Wrap("Visible text"));
        using var cancellation = new CancellationTokenSource();
        var context = new LatexProjectionContext(document, cancellation.Token);
        IEnumerable<LatexSourceSpan> Represented() {
            yield return document.Source.CreateSpan(0, 1);
            cancellation.Cancel();
            yield return document.Source.CreateSpan(1, 2);
        }
        Assert.Throws<OperationCanceledException>(() => context.ExtractResidual(document.SyntaxTree.Root.Span, Represented()));
    }
}

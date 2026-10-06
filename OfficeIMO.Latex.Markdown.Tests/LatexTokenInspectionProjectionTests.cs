namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexTokenInspectionProjectionTests {
    [Fact]
    public void CompletePublicInspectionKeepsProjectionCommentsFootnotesAndEditedSource() {
        string source = "\\documentclass{article}\\begin{document}\\section{Title}" +
            string.Concat(Enumerable.Repeat("Visible % hidden\ntext \\textbf{bold}.\n\n", 300)) +
            "End\\footnote{Note body.}\\end{document}";
        LatexDocument document = LatexDocument.Parse(source);
        LatexToMarkdownResult expected = document.ToMarkdownDocumentResult();
        foreach (LatexToken token in document.Tokens) {
            Assert.Equal(source.Substring(token.Span.Start.Offset, token.Span.Length), token.Text);
            _ = token.Value;
        }

        LatexToMarkdownResult actual = document.ToMarkdownDocumentResult();
        Assert.Equal(expected.Value.ToMarkdown(), actual.Value.ToMarkdown());
        Assert.Equal(expected.Report.Diagnostics.Select(diagnostic => diagnostic.Code),
            actual.Report.Diagnostics.Select(diagnostic => diagnostic.Code));
        Assert.DoesNotContain("hidden", actual.Value.ToMarkdown());
        Assert.Contains("Note body.", actual.Value.ToMarkdown());
        Assert.Equal(source, document.ToLatex());
        document.Headings.Single().Title = "Edited";
        Assert.Contains("Edited", document.ToMarkdownDocument().ToMarkdown());
        Assert.Equal(source.Replace("{Title}", "{Edited}"), document.ToLatex());
    }
}

namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexTokenInspectionProjectionTests {
    [Fact]
    public void CompletePublicInspectionKeepsProjectionCommentsFootnotesAndEditedSource() {
        string source = "\\documentclass{article}\\begin{document}\\section{Title}" +
            string.Concat(Enumerable.Repeat("Visible % hidden\ntext \\textbf{bold}.\n\n", 300)) +
            "End\\footnote{Note body.}\\end{document}";
        LatexDocument document = LatexDocument.Parse(source);
        var expected = document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport1);
        var expectedReport = (LatexToMarkdownReport)conversionReport1;
        foreach (LatexToken token in document.Tokens) {
            Assert.Equal(source.Substring(token.Span.Start.Offset, token.Span.Length), token.Text);
            _ = token.Value;
        }

        var actual = document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport2);
        var actualReport = (LatexToMarkdownReport)conversionReport2;
        Assert.Equal(expected.ToMarkdown(), actual.ToMarkdown());
        Assert.Equal(expectedReport.Diagnostics.Select(diagnostic => diagnostic.Code),
            actualReport.Diagnostics.Select(diagnostic => diagnostic.Code));
        Assert.DoesNotContain("hidden", ((actual.ToMarkdown())));
        Assert.Contains("Note body.", ((actual.ToMarkdown())));
        Assert.Equal(source, document.ToLatex());
        document.Headings.Single().Title = "Edited";
        Assert.Contains("Edited", ((document.ToMarkdown().ToMarkdown())));
        Assert.Equal(source.Replace("{Title}", "{Edited}"), document.ToLatex());
    }
}

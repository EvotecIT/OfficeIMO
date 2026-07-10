namespace OfficeIMO.Latex.Tests;

public sealed class LatexWriterTests {
    [Fact]
    public void SemanticEdits_ReplaceOnlyArgumentAndMathContentSpans() {
        const string source = "\\documentclass{article}\r\n\\begin{document}\r\n\\section[Old]{Old title}\r\nValue $x+1$.\r\n\\end{document}\r\n";
        LatexDocument document = LatexDocument.Parse(source).Document;

        LatexHeading heading = Assert.Single(document.Headings);
        heading.Title = "New title";
        heading.ShortTitle = "New";
        Assert.Single(document.Math).Content = "y^2";

        Assert.True(document.IsModified);
        Assert.Equal("\\documentclass{article}\r\n\\begin{document}\r\n\\section[New]{New title}\r\nValue $y^2$.\r\n\\end{document}\r\n", document.ToLatex());
    }

    [Fact]
    public void ParagraphEdit_PreservesSurroundingCommandsAndLineEndings() {
        const string source = "\\documentclass{article}\n\\begin{document}\n\\section{One}\nOld paragraph.\n\n\\section{Two}\nOther.\n\\end{document}";
        LatexDocument document = LatexDocument.Parse(source).Document;
        LatexParagraph paragraph = Assert.Single(document.Paragraphs, item => item.Content == "Old paragraph.");

        paragraph.Content = "New paragraph with \\emph{markup}.";

        Assert.Contains("\\section{One}\nNew paragraph with \\emph{markup}.\n\n\\section{Two}", document.ToLatex(), StringComparison.Ordinal);
    }

    [Fact]
    public void CanonicalMode_NormalizesMixedLineEndingsOnly() {
        const string source = "\\documentclass{article}\r\n\\begin{document}\rText\n\\end{document}";
        LatexDocument document = LatexDocument.Parse(source).Document;

        string canonical = document.ToLatex(new LatexWriterOptions { Mode = LatexWriterMode.Canonical, LineEnding = "\n" });

        Assert.Equal("\\documentclass{article}\n\\begin{document}\nText\n\\end{document}", canonical);
    }

    [Fact]
    public void EditingContainerAndNestedArgumentTogether_IsRejected() {
        const string source = "\\documentclass{article}\n\\begin{document}\n\\textbf{old}\n\\end{document}";
        LatexDocument document = LatexDocument.Parse(source).Document;
        document.Body!.Content = "replacement";
        Assert.Single(document.Commands, command => command.Name == "textbf").GetRequiredArgument(0)!.Content = "new";

        Assert.Throws<InvalidOperationException>(() => document.ToLatex());
    }
}

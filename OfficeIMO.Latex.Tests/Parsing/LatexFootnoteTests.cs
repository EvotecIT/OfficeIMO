namespace OfficeIMO.Latex.Tests;

public sealed class LatexFootnoteTests {
    [Fact]
    public void Footnotes_retain_source_order_marks_and_editable_bodies() {
        const string source = @"\begin{document}Text\footnote[7]{First \emph{note}.} and\footnote X.\end{document}";
        LatexDocument document = LatexDocument.Parse(source);
        Assert.Equal(2, document.Footnotes.Count);
        Assert.Equal("7", document.Footnotes[0].Mark);
        Assert.Equal(@"First \emph{note}.", document.Footnotes[0].Content);
        Assert.Equal("X", document.Footnotes[1].Content);
        Assert.True(document.Footnotes[1].Body.IsSingleToken);
        Assert.Equal(source, document.ToLatex());

        document.Footnotes[0].Mark = "9";
        document.Footnotes[1].Content = "Second note";
        string written = document.ToLatex();
        Assert.Equal(@"\begin{document}Text\footnote[9]{First \emph{note}.} and\footnote {Second note}.\end{document}", written);
        LatexDocument reopened = LatexDocument.Parse(written);
        Assert.Equal("9", reopened.Footnotes[0].Mark);
        Assert.Equal("Second note", reopened.Footnotes[1].Content);
        Assert.True(reopened.SyntaxTree.IsLossless);
    }

    [Fact]
    public void Definition_body_and_preserve_only_do_not_create_active_footnotes() {
        const string source = @"\newcommand{\saved}{\footnote{Inert}}\begin{document}Body\footnote{Visible}\end{document}";
        Assert.Equal("Visible", Assert.Single(LatexDocument.Parse(source).Footnotes).Content);
        LatexDocument preserved = LatexDocument.Parse(source, LatexParseOptions.CreateProfile(LatexDocumentProfile.PreserveOnly));
        Assert.Empty(preserved.Footnotes);
        Assert.Equal(source, preserved.ToLatex());
    }

    [Fact]
    public void Footnote_and_argument_edits_share_the_same_conflict_checked_source_owner() {
        LatexDocument document = LatexDocument.Parse(@"\footnote{Old}");
        LatexFootnote footnote = Assert.Single(document.Footnotes);
        footnote.Content = "New";
        Assert.Equal(@"\footnote{New}", document.ToLatex());
        footnote.Command.GetRequiredArgument(0)!.Content = "Latest";
        Assert.Equal("Latest", footnote.Content);
        Assert.Equal(@"\footnote{Latest}", document.ToLatex());
    }
}

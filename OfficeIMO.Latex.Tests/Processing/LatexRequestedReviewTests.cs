namespace OfficeIMO.Latex.Tests;

public sealed class LatexRequestedReviewTests {
    [Fact]
    public void SafeMacroMayContainAnExplicitNavigationLink() {
        var document = LatexDocument.Parse(@"\newcommand{\jump}{\hyperref[target]{Visible}}",
            new LatexParseOptions { MacroExpansion = LatexMacroExpansion.SafeSimpleDefinitions });
        Assert.True(Assert.Single(document.MacroDefinitions).IsSafe);
        Assert.Equal(@"\hyperref[target]{Visible}", document.ExpandSimpleMacros(@"\jump").Value);
    }

    [Theory]
    [InlineData("{A]B}")]
    [InlineData("{A[B]C}")]
    [InlineData("{A% ]\nB}")]
    public void OptionalMacroArgumentsRetainBraceProtectedBrackets(string argument) {
        var document = LatexDocument.Parse(@"\newcommand{\f}[1][D]{<#1>}",
            new LatexParseOptions { MacroExpansion = LatexMacroExpansion.SafeSimpleDefinitions });
        Assert.Equal("<" + argument + ">", document.ExpandSimpleMacros("\\f[" + argument + "]").Value);
    }

    [Fact]
    public void MacroExpansionRebindsEditedDefaultsAndTransitiveSafety() {
        var document = LatexDocument.Parse(@"\newcommand{\term}[1][old]{#1}\newcommand{\wrapper}{\term}",
            new LatexParseOptions { MacroExpansion = LatexMacroExpansion.SafeSimpleDefinitions });
        document.MacroDefinitions[0].Command.Arguments.Single(argument => argument.IsOptional && argument.Content == "old").Content = "new";
        Assert.Equal("new", document.ExpandSimpleMacros(@"\wrapper").Value);
        document.MacroDefinitions[0].Command.GetRequiredArgument(1)!.Content = @"\input{private}";
        Assert.Equal(@"\wrapper", document.ExpandSimpleMacros(@"\wrapper").Value);
    }

    [Fact]
    public void SafetyClassificationPropagatesAcrossLongChainsWithoutChangingCyclesOrSafeLeaves() {
        var source = new System.Text.StringBuilder(@"\newcommand{\cyclea}{\cycleb}\newcommand{\cycleb}{\cyclea}\newcommand{\safe}{\textbf{visible}}");
        string Name(int index) => "chain" + (char)('a' + index / 26) + (char)('a' + index % 26);
        for (int index = 0; index < 400; index++) source.Append("\\newcommand{\\" + Name(index) + "}{" +
            (index < 399 ? "\\" + Name(index + 1) : "\\input{private}") + "}");
        var document = LatexDocument.Parse(source.ToString());
        Assert.Equal(403, document.MacroDefinitions.Count);
        Assert.All(document.MacroDefinitions.Where(definition => definition.Name.StartsWith("chain", StringComparison.Ordinal)),
            definition => Assert.False(definition.IsSafe));
        Assert.All(document.MacroDefinitions.Take(3), definition => Assert.True(definition.IsSafe));
    }

    [Theory]
    [InlineData(@"TRAILING\section{Inactive}")]
    [InlineData(@"TRAILING\[inactive\]")]
    public void ParagraphSpansRemainInsideTheActiveDocumentBody(string trailer) {
        string source = @"\documentclass{article}\begin{document}Public\end{document}" + trailer;
        var document = LatexDocument.Parse(source);
        Assert.Equal(source, document.ToLatex());
        Assert.Equal("Public", Assert.Single(document.Paragraphs).Content);
        Assert.All(document.Paragraphs, paragraph => {
            Assert.True(paragraph.Span.Start.Offset >= document.Body!.ContentSpan.Start.Offset);
            Assert.True(paragraph.Span.End.Offset <= document.Body.ContentSpan.End.Offset);
        });
    }
}

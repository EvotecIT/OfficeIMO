namespace OfficeIMO.Latex.Tests;

public sealed class LatexProfileRegressionTests {
    [Fact]
    public void PreserveOnly_RetainsStructuralModelWithoutOfficeProfileSemantics() {
        const string source =
            "\\documentclass{article}\n" +
            "\\begin{document}\n" +
            "\\section{Heading}\n" +
            "Body.\n" +
            "\\begin{itemize}\\item One\\end{itemize}\n" +
            "\\end{document}\n";

        LatexDocument document = LatexDocument.Parse(source, new LatexParseOptions {
            Profile = LatexDocumentProfile.PreserveOnly
        }).Document;

        Assert.NotEmpty(document.Commands);
        Assert.NotEmpty(document.Environments);
        Assert.Empty(document.Headings);
        Assert.Empty(document.Paragraphs);
        Assert.Empty(document.Lists);
        Assert.Empty(document.Figures);
        Assert.Empty(document.Tables);
        Assert.Empty(document.Citations);
        Assert.Empty(document.References);
        Assert.Empty(document.Labels);
        Assert.Empty(document.Theorems);
        Assert.Empty(document.MacroDefinitions);
        Assert.False(document.IsRecognizedProfile);
        Assert.Equal(source, document.ToLatex());
    }

    [Fact]
    public void ProfileSignatures_KeepBracedItemContentAndPositionedTabularArguments() {
        const string source =
            "\\documentclass{article}\n" +
            "\\begin{document}\n" +
            "\\begin{itemize}\\item {Visible item}\\end{itemize}\n" +
            "\\begin{tabular}[t]{ll}A & B\\\\\\end{tabular}\n" +
            "\\end{document}\n";

        LatexDocument document = LatexDocument.Parse(source).Document;

        LatexListItem item = Assert.Single(Assert.Single(document.Lists).Items);
        Assert.Equal("{Visible item}", item.Content);
        Assert.DoesNotContain(item.ItemCommand.Arguments, static argument => !argument.IsOptional);
        LatexTable table = Assert.Single(document.Tables);
        Assert.Equal("t", table.Environment.BeginCommand.GetOptionalArgument(0)?.Content);
        Assert.Equal("ll", table.ColumnSpecification);
        Assert.Equal(source, document.ToLatex());
    }

    [Fact]
    public void StarredHeading_BindsItsTitleWithoutLosingTheModifier() {
        const string source = "\\documentclass{article}\n\\begin{document}\n\\section*{Unnumbered}\n\\end{document}\n";

        LatexDocument document = LatexDocument.Parse(source).Document;

        LatexHeading heading = Assert.Single(document.Headings);
        Assert.Equal("Unnumbered", heading.Title);
        Assert.True(heading.Command.IsStarred);
        Assert.Equal(source, document.ToLatex());
    }

    [Fact]
    public void ZeroArgumentCommand_DoesNotClaimFollowingStandaloneGroup() {
        const string source = "\\documentclass{article}\n\\begin{document}\n\\maketitle\n{Visible group}\n\\end{document}\n";

        LatexDocument document = LatexDocument.Parse(source).Document;

        LatexCommand makeTitle = Assert.Single(document.Commands, static command => command.Name == "maketitle");
        Assert.Empty(makeTitle.Arguments);
        Assert.Contains(document.Paragraphs, static paragraph => paragraph.Content == "{Visible group}");
        Assert.Equal(source, document.ToLatex());
    }

    [Fact]
    public void LiteralSquareBrackets_AreTextUnlessACommandSignatureClaimsThem() {
        const string source = "\\documentclass{article}\n\\begin{document}\nText [literal] and [unterminated.\n\\end{document}\n";

        LatexParseResult result = LatexDocument.Parse(source);

        Assert.DoesNotContain(result.Document.SyntaxTree.Root.DescendantsAndSelf(),
            static node => node.Kind == LatexSyntaxKind.OptionalGroup);
        Assert.DoesNotContain(result.Diagnostics, static diagnostic => diagnostic.Code is "LATEX001" or "LATEX002");
        Assert.Contains(result.Document.Paragraphs,
            static paragraph => paragraph.Content.IndexOf("[literal] and [unterminated", StringComparison.Ordinal) >= 0);
        Assert.Equal(source, result.Document.ToLatex());
    }
}

namespace OfficeIMO.Latex.Tests;

public sealed class LatexSingleTokenArgumentTests {
    [Fact]
    public void Starred_macro_definition_after_control_word_trivia_keeps_name_and_body() {
        const string source = "\\newcommand % comment\n*\\x{Body}";
        LatexDocument document = LatexDocument.Parse(source);
        LatexMacroDefinition definition = Assert.Single(document.MacroDefinitions);
        Assert.True(definition.Command.IsStarred);
        Assert.Equal("x", definition.Name);
        Assert.Equal("Body", definition.Body);
        Assert.DoesNotContain(document.Diagnostics, diagnostic => diagnostic.Code == "LATEX007");
        Assert.Equal(source, document.ToLatex());
        Assert.True(document.SyntaxTree.IsLossless);
    }

    [Theory]
    [InlineData(@"\section*XYZ", "section", "X")]
    [InlineData(@"\part*😀tail", "part", "😀")]
    [InlineData(@"\includegraphics*XYZ", "includegraphics", "X")]
    [InlineData("\\section % comment\n*XYZ", "section", "X")]
    public void Star_modifier_precedes_single_token_argument_and_is_preserved_when_editing(string source, string name, string value) {
        LatexDocument document = LatexDocument.Parse(source);
        LatexCommand command = Assert.Single(document.Commands, item => item.Name == name);
        Assert.True(command.IsStarred);
        LatexArgument argument = command.GetRequiredArgument(0)!;
        Assert.Equal(value, argument.Content);
        Assert.Equal(source, document.ToLatex());
        Assert.True(document.SyntaxTree.IsLossless);
        Assert.Equal(source, string.Concat(document.Tokens.Select(token => token.Text)));
        argument.Content = "New title";
        string expected = source.Substring(0, argument.ContentSpan.Start.Offset) + "{New title}" + source.Substring(argument.ContentSpan.End.Offset);
        Assert.Equal(expected, document.ToLatex());
        Assert.True(Assert.Single(LatexDocument.Parse(expected).Commands, item => item.Name == name).IsStarred);
    }

    [Theory]
    [InlineData(@"\textbf XYZ", "X")]
    [InlineData(@"\textbf 😀tail", "😀")]
    [InlineData(@"\textbf\% tail", @"\%")]
    [InlineData("\\textbf% delimiter\n XYZ", "X")]
    public void Required_argument_binds_one_token_without_changing_lossless_inventory(string source, string expected) {
        LatexDocument document = LatexDocument.Parse(source);
        LatexArgument argument = Assert.Single(document.Commands, command => command.Name == "textbf").GetRequiredArgument(0)!;
        Assert.NotNull(argument);
        Assert.True(argument.IsSingleToken);
        Assert.True(argument.IsTerminated);
        Assert.Equal(expected, argument.Content);
        Assert.Equal(expected, argument.ContentSpan.Slice(source));
        Assert.Equal(source, string.Concat(document.Tokens.Select(token => token.Text)));
        Assert.Equal(source, document.ToLatex());
        Assert.True(document.SyntaxTree.IsLossless);
        Assert.DoesNotContain(document.Diagnostics, diagnostic => diagnostic.Code == "LATEX007");
    }

    [Fact]
    public void Multiple_required_arguments_bind_separate_characters_from_one_coalesced_token() {
        const string source = @"\href ABtail";
        LatexDocument document = LatexDocument.Parse(source);
        LatexCommand link = Assert.Single(document.Commands);
        Assert.Equal("A", link.GetRequiredArgument(0)?.Content);
        Assert.Equal("B", link.GetRequiredArgument(1)?.Content);
        Assert.Equal(@"\href AB", link.Syntax.Span.Slice(source));
        Assert.Equal(source, document.ToLatex());
        Assert.True(document.SyntaxTree.IsLossless);
    }

    [Fact]
    public void Editing_an_unbraced_argument_groups_replacement_and_preserves_following_text() {
        LatexDocument document = LatexDocument.Parse(@"\textbf XYZ");
        document.Commands[0].GetRequiredArgument(0)!.Content = "New words";
        string edited = document.ToLatex();
        Assert.Equal(@"\textbf {New words}YZ", edited);
        LatexDocument reopened = LatexDocument.Parse(edited);
        Assert.Equal("New words", reopened.Commands[0].GetRequiredArgument(0)?.Content);
        Assert.True(reopened.SyntaxTree.IsLossless);
    }

    [Fact]
    public void Captured_control_word_delimiter_belongs_to_command_but_not_argument() {
        const string source = "\\textbf\\textasciitilde % comment\r\n next";
        LatexDocument document = LatexDocument.Parse(source);
        LatexCommand bold = Assert.Single(document.Commands, command => command.Name == "textbf");
        Assert.Equal(@"\textasciitilde", bold.GetRequiredArgument(0)?.Content);
        Assert.Equal(source.IndexOf("next", StringComparison.Ordinal), bold.Syntax.Span.End.Offset);
        Assert.Equal(source, document.ToLatex());
        Assert.True(document.SyntaxTree.IsLossless);
    }

    [Fact]
    public void Missing_argument_does_not_capture_a_closing_group_delimiter() {
        LatexDocument document = LatexDocument.Parse(@"{\textbf }");
        Assert.Null(Assert.Single(document.Commands).GetRequiredArgument(0));
        Assert.Contains(document.Diagnostics, diagnostic => diagnostic.Code == "LATEX007");
        Assert.True(document.SyntaxTree.IsLossless);
    }

    [Theory]
    [InlineData("\\textbf\n\nX")]
    [InlineData("\\textbf\n\n{X}")]
    public void Argument_binding_does_not_skip_a_paragraph_boundary(string source) {
        LatexDocument document = LatexDocument.Parse(source);
        Assert.Null(Assert.Single(document.Commands).GetRequiredArgument(0));
        Assert.Contains(document.Diagnostics, diagnostic => diagnostic.Code == "LATEX007");
        Assert.Equal(source, document.ToLatex());
        Assert.True(document.SyntaxTree.IsLossless);
    }
}

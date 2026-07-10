namespace OfficeIMO.Latex.Tests;

public sealed class LatexTechnicalSemanticsTests {
    private const string TechnicalDocument =
        "\\documentclass{article}\n" +
        "\\newcommand{\\term}[2][Default]{\\textbf{#1}: #2}\n" +
        "\\newcommand{\\danger}[1]{\\input{#1}}\n" +
        "\\begin{document}\n" +
        "\\section{Results}\\label{sec:results}\n" +
        "See \\ref{sec:results} and \\citep[see]{key1,key2}.\n\n" +
        "\\begin{itemize}\n" +
        "\\item First\n" +
        "\\item[Named] Second\n" +
        "\\end{itemize}\n\n" +
        "\\begin{figure}[ht]\n" +
        "\\includegraphics[width=.5\\textwidth]{plot.png}\n" +
        "\\caption{A plot}\n" +
        "\\label{fig:plot}\n" +
        "\\end{figure}\n\n" +
        "\\begin{table}\n" +
        "\\caption{Values}\n" +
        "\\begin{tabular}{lc}\n" +
        "Name & Value\\\\\n" +
        "A & \\textbf{B}\\\\\n" +
        "\\end{tabular}\n" +
        "\\end{table}\n\n" +
        "\\begin{theorem}[Pythagoras]\\label{thm:p}\n" +
        "For a right triangle, $a^2+b^2=c^2$.\n" +
        "\\end{theorem}\n" +
        "\\end{document}\n";

    [Fact]
    public void TechnicalDocument_BindsListsFiguresTablesCitationsReferencesAndTheorems() {
        LatexDocument document = LatexDocument.Parse(TechnicalDocument).Document;

        LatexList list = Assert.Single(document.Lists);
        Assert.Equal(LatexListKind.Unordered, list.Kind);
        Assert.Equal(2, list.Items.Count);
        Assert.Equal("First", list.Items[0].Content);
        Assert.Equal("Named", list.Items[1].Label);
        Assert.Equal("Second", list.Items[1].Content);

        LatexFigure figure = Assert.Single(document.Figures);
        Assert.Equal("plot.png", Assert.Single(figure.Images).Target);
        Assert.Equal("width=.5\\textwidth", figure.Images[0].Options);
        Assert.Equal("A plot", figure.Caption);
        Assert.Equal("fig:plot", figure.Label);

        LatexTable table = Assert.Single(document.Tables);
        Assert.Equal("lc", table.ColumnSpecification);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(new[] { "Name", "Value" }, table.Rows[0].Cells.Select(cell => cell.Content));
        Assert.Equal("\\textbf{B}", table.Rows[1].Cells[1].Content);

        LatexCitation citation = Assert.Single(document.Citations);
        Assert.Equal(new[] { "key1", "key2" }, citation.Keys);
        Assert.Equal("see", citation.Prenote);
        Assert.Contains(document.References, reference => reference.Target == "sec:results");
        Assert.Contains(document.Labels, label => label.Name == "thm:p");

        LatexTheorem theorem = Assert.Single(document.Theorems);
        Assert.Equal("Pythagoras", theorem.Title);
        Assert.Equal("thm:p", theorem.Label);
        Assert.Contains("a^2+b^2=c^2", theorem.Content, StringComparison.Ordinal);
        Assert.Equal(TechnicalDocument, document.ToLatex());
    }

    [Fact]
    public void TechnicalSemanticEdits_ReplaceOnlyOwnedSourceSlices() {
        LatexDocument document = LatexDocument.Parse(TechnicalDocument).Document;

        document.Lists[0].Items[0].Content = "Updated first";
        document.Figures[0].Images[0].Target = "updated.pdf";
        document.Figures[0].Caption = "Updated plot";
        document.Tables[0].Rows[1].Cells[1].Content = "\\emph{Changed}";

        string updated = document.ToLatex();
        Assert.Contains("\\item Updated first", updated, StringComparison.Ordinal);
        Assert.Contains("{updated.pdf}", updated, StringComparison.Ordinal);
        Assert.Contains("\\caption{Updated plot}", updated, StringComparison.Ordinal);
        Assert.Contains("A & \\emph{Changed}\\\\", updated, StringComparison.Ordinal);
        Assert.Contains("\\begin{theorem}[Pythagoras]", updated, StringComparison.Ordinal);
    }

    [Fact]
    public void SimpleMacroDefinitions_AreClassifiedAndExpandedOnlyWhenExplicitlyEnabled() {
        var options = new LatexParseOptions { MacroExpansion = LatexMacroExpansion.SafeSimpleDefinitions };
        LatexDocument document = LatexDocument.Parse(TechnicalDocument, options).Document;

        LatexMacroDefinition safe = Assert.Single(document.MacroDefinitions, definition => definition.Name == "term");
        LatexMacroDefinition unsafeDefinition = Assert.Single(document.MacroDefinitions, definition => definition.Name == "danger");
        Assert.True(safe.IsSafe);
        Assert.Equal(2, safe.ParameterCount);
        Assert.Equal("Default", safe.DefaultValue);
        Assert.False(unsafeDefinition.IsSafe);

        LatexMacroExpansionResult explicitValue = document.ExpandSimpleMacros("\\term[Name]{Value}");
        LatexMacroExpansionResult defaultValue = document.ExpandSimpleMacros("\\term{Value}");
        Assert.Equal("\\textbf{Name}: Value", explicitValue.Value);
        Assert.Equal("\\textbf{Default}: Value", defaultValue.Value);
        Assert.Empty(explicitValue.Diagnostics);
        Assert.Equal(TechnicalDocument, document.ToLatex());
    }

    [Fact]
    public void MacroExpansion_RemainsDisabledUnlessRequested() {
        LatexDocument document = LatexDocument.Parse(TechnicalDocument).Document;

        Assert.Throws<InvalidOperationException>(() => document.ExpandSimpleMacros("\\term{Value}"));
    }

    [Fact]
    public void CyclicSafeMacros_AreDiagnosedAndBounded() {
        const string source = "\\newcommand{\\a}{\\b}\\newcommand{\\b}{\\a}";
        LatexDocument document = LatexDocument.Parse(source, new LatexParseOptions {
            MacroExpansion = LatexMacroExpansion.SafeSimpleDefinitions
        }).Document;

        LatexMacroExpansionResult result = document.ExpandSimpleMacros("\\a");

        Assert.Equal("\\a", result.Value);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "LATEXMAC002");
    }
}

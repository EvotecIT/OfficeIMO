namespace OfficeIMO.Latex;

/// <summary>Lossless, non-executing LaTeX parser.</summary>
internal static class LatexParser {
    /// <summary>Parses tokens, nested syntax, and the bounded OfficeIMO profile.</summary>
    public static LatexParseResult Parse(string source, LatexParseOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        options ??= new LatexParseOptions();
        IReadOnlyList<LatexToken> tokens = LatexTokenizer.Tokenize(source, options);
        var sourceText = new LatexSourceText(source);
        var diagnostics = new List<LatexDiagnostic>();
        var structural = new LatexStructuralParser(sourceText, tokens, options, diagnostics);
        LatexSyntaxTree syntaxTree = structural.Parse();
        if (!syntaxTree.IsLossless) {
            diagnostics.Add(new LatexDiagnostic(
                "LATEX900",
                LatexDiagnosticSeverity.Error,
                "Parser did not retain contiguous complete source coverage.",
                syntaxTree.Root.Span));
        }
        var document = new LatexDocument(sourceText, syntaxTree, tokens, diagnostics, options);
        return new LatexParseResult(document, diagnostics);
    }
}

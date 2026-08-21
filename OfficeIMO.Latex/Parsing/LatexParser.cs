using System.Threading;

namespace OfficeIMO.Latex;

/// <summary>Lossless, non-executing LaTeX parser.</summary>
internal static class LatexParser {
    /// <summary>Parses tokens, nested syntax, and the bounded OfficeIMO profile.</summary>
    public static LatexParseResult Parse(
        string source,
        LatexParseOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        options ??= new LatexParseOptions();
        options.ValidateNamedModes();
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<LatexToken> tokens = LatexTokenizer.Tokenize(source, options, cancellationToken);
        var sourceText = new LatexSourceText(source);
        var diagnostics = new List<LatexDiagnostic>();
        var structural = new LatexStructuralParser(sourceText, tokens, options, diagnostics, cancellationToken);
        LatexSyntaxTree syntaxTree = structural.Parse();
        if (!syntaxTree.IsLossless) {
            diagnostics.Add(new LatexDiagnostic(
                "LATEX900",
                LatexDiagnosticSeverity.Error,
                "Parser did not retain contiguous complete source coverage.",
                syntaxTree.Root.Span));
        }
        cancellationToken.ThrowIfCancellationRequested();
        var document = new LatexDocument(sourceText, syntaxTree, tokens, diagnostics, options, cancellationToken);
        return new LatexParseResult(document, diagnostics);
    }
}

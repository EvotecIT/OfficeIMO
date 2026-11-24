using System.Net;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown;

namespace OfficeIMO.Notebook.Html;

internal static class NotebookHtmlMath {
    // Keep individual expressions bounded independently of the notebook and Markdown-cell limits.
    private const int MaxExpressionCharacters = 32 * 1024;

    internal static void Register(MarkdownHtmlExportOptions options, List<OfficeDiagnostic> diagnostics, CancellationToken cancellationToken) {
        options.BlockRenderExtensions.Add(new MarkdownBlockHtmlRenderExtension("NotebookMathBlock", typeof(MathBlock),
            (block, _) => Render(((MathBlock)block).Content, true, diagnostics, cancellationToken)));
        options.InlineRenderExtensions.Add(new MarkdownInlineHtmlRenderExtension("NotebookMathInline", typeof(MathInline),
            (inline, _) => Render(((MathInline)inline).Source, ((MathInline)inline).IsDisplay, diagnostics, cancellationToken)));
        options.SemanticFencedBlockHtmlRenderer = (block, _) =>
            string.Equals(block.SemanticKind, MarkdownSemanticKinds.Math, StringComparison.OrdinalIgnoreCase)
            && !string.Equals(block.Language, "asciimath", StringComparison.OrdinalIgnoreCase)
                ? Render(block.Content, true, diagnostics, cancellationToken)
                : null;
    }

    private static string Render(string tex, bool display, List<OfficeDiagnostic> diagnostics, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (tex.Length <= MaxExpressionCharacters) {
            try {
                return OfficeMathMarkup.ToMathMl(OfficeMathMarkup.FromLatex(tex), display);
            } catch (Exception ex) when (ex is FormatException or ArgumentException or NotSupportedException) {
                // Unsupported math remains visible without executable HTML or a remote renderer.
            }
        }
        diagnostics.Add(new OfficeDiagnostic("notebook-html-math-unsupported", OfficeDiagnosticSeverity.Warning,
            "The expression was displayed as source because the bounded math renderer could not represent it.",
            format: "Jupyter", lossKind: OfficeConversionLossKind.Approximation));
        string tag = display ? "pre" : "code";
        return "<" + tag + ">" + WebUtility.HtmlEncode(tex) + "</" + tag + ">";
    }
}

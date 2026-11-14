using System.Runtime.CompilerServices;

namespace OfficeIMO.Markdown.Benchmarks;

/// <summary>Public parse routes for the semantic-only and syntax-tree benchmark lanes.</summary>
internal static class MarkdownBenchmarkParsing {
    private static readonly ConditionalWeakTable<MarkdownLoadOptions, MarkdownLoadOptions> SemanticOptions = new();

    /// <summary>Parses into the typed document without capturing the syntax tree (<see cref="MarkdownLoadOptions.CaptureSyntaxTree"/> false).</summary>
    internal static MarkdownDocument ParseSemantic(string markdown, MarkdownLoadOptions options) =>
        MarkdownDocument.Parse(markdown, SemanticOptions.GetValue(options, static source => {
            MarkdownLoadOptions semantic = source.Clone();
            semantic.CaptureSyntaxTree = false;
            return semantic;
        }));

    /// <summary>Parses with the syntax tree, transform diagnostics and source snapshot attached.</summary>
    internal static MarkdownParseInfo ParseWithSyntaxTree(string markdown, MarkdownLoadOptions? options = null) =>
        MarkdownDocument.Parse(markdown, options).ParseInfo
        ?? throw new InvalidOperationException("Markdown parse did not attach its syntax snapshot.");
}

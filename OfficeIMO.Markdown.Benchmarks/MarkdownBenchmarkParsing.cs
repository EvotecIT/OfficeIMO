using System.Runtime.CompilerServices;

namespace OfficeIMO.Markdown.Benchmarks;

/// <summary>Public parse routes for the semantic-only and syntax-tree benchmark lanes.</summary>
internal static class MarkdownBenchmarkParsing {
    private static readonly ConditionalWeakTable<MarkdownLoadOptions, MarkdownLoadOptions> SemanticOptions = new();
    private static readonly ConditionalWeakTable<MarkdownLoadOptions, MarkdownLoadOptions> SourceBackedOptions = new();
    private static readonly MarkdownLoadOptions DefaultSourceBackedOptions = new() { CaptureSyntaxTree = true };

    /// <summary>Parses into the typed document without capturing the syntax tree (<see cref="MarkdownLoadOptions.CaptureSyntaxTree"/> false).</summary>
    internal static MarkdownDocument ParseSemantic(string markdown, MarkdownLoadOptions options) =>
        MarkdownDocument.Parse(markdown, SemanticOptions.GetValue(options, static source => {
            MarkdownLoadOptions semantic = source.Clone();
            semantic.CaptureSyntaxTree = false;
            return semantic;
        }));

    /// <summary>Parses with <see cref="MarkdownLoadOptions.CaptureSyntaxTree"/> turned on.</summary>
    internal static MarkdownDocument ParseSourceBacked(string markdown, MarkdownLoadOptions? options = null) =>
        MarkdownDocument.Parse(markdown, options == null
            ? DefaultSourceBackedOptions
            : SourceBackedOptions.GetValue(options, static source => {
                MarkdownLoadOptions sourceBacked = source.Clone();
                sourceBacked.CaptureSyntaxTree = true;
                return sourceBacked;
            }));

    /// <summary>Parses with the syntax tree, transform diagnostics and source snapshot attached.</summary>
    internal static MarkdownParseInfo ParseWithSyntaxTree(string markdown, MarkdownLoadOptions? options = null) =>
        ParseSourceBacked(markdown, options).ParseInfo
        ?? throw new InvalidOperationException("Markdown parse did not attach its syntax snapshot.");
}

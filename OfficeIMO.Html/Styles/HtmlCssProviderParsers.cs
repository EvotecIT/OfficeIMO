using AngleSharp.Css.Parser;

namespace OfficeIMO.Html;

/// <summary>
/// Reuses the provider CSS parsers on each thread. Constructing a parser builds a new browsing
/// context with its CSS configuration and declaration factories, which cost more than parsing a
/// shorthand value, and the style engine parses one for every font, background and grid
/// declaration. A parser keeps no state between parses; each thread keeps its own because
/// provider events and contexts are not documented as thread-safe.
/// </summary>
internal static class HtmlCssProviderParsers {
    [ThreadStatic] private static CssParser? _declarations;
    [ThreadStatic] private static CssParser? _stylesheets;

    /// <summary>The parser with default options, for declarations and stylesheets that drop unknown declarations.</summary>
    internal static CssParser Declarations => _declarations ??= new CssParser();

    /// <summary>The parser that keeps unknown declarations, for author stylesheet rules.</summary>
    internal static CssParser StylesheetsWithUnknownDeclarations => _stylesheets ??= new CssParser(new CssParserOptions {
        IsIncludingUnknownDeclarations = true
    });
}

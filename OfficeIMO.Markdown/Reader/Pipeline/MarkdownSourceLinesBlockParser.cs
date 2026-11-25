namespace OfficeIMO.Markdown;

/// <summary>
/// Built-in block parser that reads lines through <see cref="MarkdownSourceLines"/>, so lines it
/// only inspects are never copied out of the source text.
/// </summary>
internal abstract class MarkdownSourceLinesBlockParser : IMarkdownBlockParser {
    public bool TryParse(string[] lines, ref int i, MarkdownLoadOptions options, MarkdownDocument doc, MarkdownReaderState state) =>
        TryParse(MarkdownSourceLines.FromArray(lines), ref i, options, doc, state);

    public abstract bool TryParse(MarkdownSourceLines lines, ref int i, MarkdownLoadOptions options, MarkdownDocument doc, MarkdownReaderState state);

    /// <summary>
    /// Runs any parser over <paramref name="lines"/>. Parsers outside the reader get every line as
    /// a string, in the array the line source reads from, so lines they rewrite stay rewritten.
    /// </summary>
    internal static bool Run(IMarkdownBlockParser parser, MarkdownSourceLines lines, ref int i, MarkdownLoadOptions options, MarkdownDocument doc, MarkdownReaderState state) =>
        parser is MarkdownSourceLinesBlockParser sourceLinesParser
            ? sourceLinesParser.TryParse(lines, ref i, options, doc, state)
            : parser.TryParse(lines.ToArray(), ref i, options, doc, state);
}

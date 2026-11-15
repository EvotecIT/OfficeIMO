namespace OfficeIMO.Markdown;

/// <summary>
/// The first character a built-in block parser needs to see on a line before it can match.
/// The reader tries every parser in the pipeline at every block start; most of them reject the
/// line on its first non-whitespace character, so the dispatch loops check that character once
/// instead of calling into each parser.
/// </summary>
internal enum MarkdownBlockStart : byte {
    /// <summary>Any line; the parser is always tried.</summary>
    Any,
    Quote,
    CodeFence,
    MathFence,
    Image,
    ThematicBreak,
    HtmlBlock,
    Bracket,
    Abbreviation,
    OrderedList,
    UnorderedList,
    IndentedCode,
    SetextHeading,
    AtxHeading,
    ColonFence,
    TableOfContents
}

internal static partial class MarkdownReader {
    internal static MarkdownBlockStart ClassifyBlockStart(IMarkdownBlockParser parser) {
        // Exact types: a parser is classified only when its rejection rules are known here.
        Type type = parser.GetType();
        if (type == typeof(QuoteParser) || type == typeof(CalloutParser)) return MarkdownBlockStart.Quote;
        if (type == typeof(FencedCodeParser)) return MarkdownBlockStart.CodeFence;
        if (type == typeof(MathBlockParser)) return MarkdownBlockStart.MathFence;
        if (type == typeof(ImageParser)) return MarkdownBlockStart.Image;
        if (type == typeof(HrParser)) return MarkdownBlockStart.ThematicBreak;
        if (type == typeof(HtmlBlockParser)) return MarkdownBlockStart.HtmlBlock;
        if (type == typeof(ReferenceLinkDefParser) || type == typeof(FootnoteParser)) return MarkdownBlockStart.Bracket;
        if (type == typeof(AbbreviationDefParser)) return MarkdownBlockStart.Abbreviation;
        if (type == typeof(OrderedListParser)) return MarkdownBlockStart.OrderedList;
        if (type == typeof(UnorderedListParser)) return MarkdownBlockStart.UnorderedList;
        if (type == typeof(IndentedCodeParser)) return MarkdownBlockStart.IndentedCode;
        if (type == typeof(SetextHeadingParser)) return MarkdownBlockStart.SetextHeading;
        if (type == typeof(HeadingParser)) return MarkdownBlockStart.AtxHeading;
        if (type == typeof(CustomContainerParser)) return MarkdownBlockStart.ColonFence;
        if (type == typeof(TocParser)) return MarkdownBlockStart.TableOfContents;
        return MarkdownBlockStart.Any;
    }

    /// <summary>First character of <paramref name="line"/> that is not white space, or '\0' for a blank line.</summary>
    internal static char GetFirstNonWhiteSpaceCharacter(string? line) {
        if (line == null) return '\0';
        for (int index = 0; index < line.Length; index++) {
            char value = line[index];
            if (!char.IsWhiteSpace(value)) return value;
        }

        return '\0';
    }

    /// <summary>
    /// False only when the parser of this kind cannot match at <paramref name="index"/>: each one
    /// requires its marker to follow nothing but white space.
    /// </summary>
    internal static bool MayStartBlock(
        MarkdownBlockStart kind,
        string[] lines,
        int index,
        char first,
        MarkdownLoadOptions options) {
        switch (kind) {
            case MarkdownBlockStart.Any:
                return true;
            case MarkdownBlockStart.Quote:
                return first == '>';
            case MarkdownBlockStart.CodeFence:
                return first is '`' or '~';
            case MarkdownBlockStart.MathFence:
                return first == '$';
            case MarkdownBlockStart.Image:
                return first is '!' or '[';
            case MarkdownBlockStart.ThematicBreak:
                return first is '-' or '*' or '_';
            case MarkdownBlockStart.HtmlBlock:
                return first == '<';
            case MarkdownBlockStart.Bracket:
                return first == '[';
            case MarkdownBlockStart.Abbreviation:
                return first == '*';
            case MarkdownBlockStart.OrderedList:
                return char.IsDigit(first) || (options.ListExtras && first is >= 'a' and <= 'z' or >= 'A' and <= 'Z');
            case MarkdownBlockStart.UnorderedList:
                return first is '-' or '*' or '+';
            case MarkdownBlockStart.IndentedCode: {
                string line = lines[index];
                return line != null && line.Length > 0 && line[0] is ' ' or '\t';
            }
            case MarkdownBlockStart.SetextHeading:
                return index + 1 < lines.Length && GetFirstNonWhiteSpaceCharacter(lines[index + 1]) is '=' or '-';
            case MarkdownBlockStart.AtxHeading:
                return first == '#';
            case MarkdownBlockStart.ColonFence:
                return first == ':';
            case MarkdownBlockStart.TableOfContents:
                return first is '[' or '{' or '<';
            default:
                return true;
        }
    }
}

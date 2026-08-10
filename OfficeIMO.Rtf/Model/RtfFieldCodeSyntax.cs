using System.Collections.ObjectModel;

namespace OfficeIMO.Rtf;

/// <summary>Semantic role of one lossless RTF field-code token.</summary>
public enum RtfFieldCodeTokenKind {
    /// <summary>Whitespace between field-code tokens.</summary>
    Whitespace,
    /// <summary>The leading field keyword, such as <c>HYPERLINK</c> or <c>PAGE</c>.</summary>
    Keyword,
    /// <summary>A backslash-prefixed field switch.</summary>
    Switch,
    /// <summary>A quoted or unquoted field argument.</summary>
    Argument
}

/// <summary>One lossless token from an RTF field instruction.</summary>
public sealed class RtfFieldCodeToken {
    internal RtfFieldCodeToken(
        RtfFieldCodeTokenKind kind,
        string text,
        string value,
        int position,
        bool isQuoted,
        bool isTerminated = true) {
        Kind = kind;
        Text = text;
        Value = value;
        Position = position;
        IsQuoted = isQuoted;
        IsTerminated = isTerminated;
    }

    /// <summary>Gets the token role.</summary>
    public RtfFieldCodeTokenKind Kind { get; }
    /// <summary>Gets the exact authored token text.</summary>
    public string Text { get; }
    /// <summary>Gets the decoded token value.</summary>
    public string Value { get; }
    /// <summary>Gets the zero-based source position.</summary>
    public int Position { get; }
    /// <summary>Gets whether the argument was quoted.</summary>
    public bool IsQuoted { get; }
    /// <summary>Gets whether a quoted argument had a closing quote.</summary>
    public bool IsTerminated { get; }
}

/// <summary>Lossless typed syntax for an RTF field instruction.</summary>
public sealed class RtfFieldCodeSyntax {
    private readonly ReadOnlyCollection<RtfFieldCodeToken> _tokens;

    private RtfFieldCodeSyntax(string text, IReadOnlyList<RtfFieldCodeToken> tokens) {
        Text = text;
        _tokens = new ReadOnlyCollection<RtfFieldCodeToken>(tokens.ToArray());
        Keyword = _tokens.FirstOrDefault(static token => token.Kind == RtfFieldCodeTokenKind.Keyword)?.Value;
    }

    /// <summary>Gets the exact authored field instruction.</summary>
    public string Text { get; }
    /// <summary>Gets the decoded leading field keyword.</summary>
    public string? Keyword { get; }
    /// <summary>Gets every token, including whitespace, in authored order.</summary>
    public IReadOnlyList<RtfFieldCodeToken> Tokens => _tokens;
    /// <summary>Gets whether every quoted argument is terminated.</summary>
    public bool IsValid => _tokens.All(static token => token.IsTerminated);

    /// <summary>Parses a field instruction without discarding trivia or unknown switches.</summary>
    public static RtfFieldCodeSyntax Parse(string instruction) {
        if (instruction == null) throw new ArgumentNullException(nameof(instruction));
        var tokens = new List<RtfFieldCodeToken>();
        bool hasKeyword = false;
        int index = 0;
        while (index < instruction.Length) {
            int start = index;
            if (char.IsWhiteSpace(instruction[index])) {
                while (index < instruction.Length && char.IsWhiteSpace(instruction[index])) index++;
                string whitespace = instruction.Substring(start, index - start);
                tokens.Add(new RtfFieldCodeToken(RtfFieldCodeTokenKind.Whitespace, whitespace, whitespace, start, false));
                continue;
            }

            if (instruction[index] == '"') {
                tokens.Add(ReadQuotedArgument(instruction, ref index));
                continue;
            }

            if (instruction[index] == '\\' && index + 1 < instruction.Length) {
                index += 2;
                while (index < instruction.Length && !char.IsWhiteSpace(instruction[index]) && instruction[index] != '"') index++;
                string text = instruction.Substring(start, index - start);
                tokens.Add(new RtfFieldCodeToken(
                    RtfFieldCodeTokenKind.Switch,
                    text,
                    text.Substring(1),
                    start,
                    false));
                continue;
            }

            while (index < instruction.Length && !char.IsWhiteSpace(instruction[index])) index++;
            string authored = instruction.Substring(start, index - start);
            RtfFieldCodeTokenKind kind = hasKeyword ? RtfFieldCodeTokenKind.Argument : RtfFieldCodeTokenKind.Keyword;
            tokens.Add(new RtfFieldCodeToken(kind, authored, authored, start, false));
            hasKeyword = true;
        }
        return new RtfFieldCodeSyntax(instruction, tokens);
    }

    private static RtfFieldCodeToken ReadQuotedArgument(string instruction, ref int index) {
        int start = index++;
        var value = new System.Text.StringBuilder();
        bool terminated = false;
        while (index < instruction.Length) {
            char character = instruction[index++];
            if (character == '"') {
                terminated = true;
                break;
            }
            if (character == '\\' && index < instruction.Length &&
                (instruction[index] == '"' || instruction[index] == '\\')) {
                value.Append(instruction[index++]);
                continue;
            }
            value.Append(character);
        }
        return new RtfFieldCodeToken(
            RtfFieldCodeTokenKind.Argument,
            instruction.Substring(start, index - start),
            value.ToString(),
            start,
            true,
            terminated);
    }
}

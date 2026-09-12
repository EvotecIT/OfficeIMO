using System;

namespace OfficeIMO.Html.Css;

/// <summary>CSS Syntax token categories, with comments retained as source trivia.</summary>
public enum HtmlCssTokenKind {
    /// <summary>An identifier.</summary>
    Identifier,
    /// <summary>An identifier and opening function parenthesis.</summary>
    Function,
    /// <summary>An at-keyword.</summary>
    AtKeyword,
    /// <summary>A hash token.</summary>
    Hash,
    /// <summary>A quoted string, including one terminated by end of input.</summary>
    String,
    /// <summary>A string terminated by an unescaped newline.</summary>
    BadString,
    /// <summary>An unquoted URL.</summary>
    Url,
    /// <summary>A malformed unquoted URL.</summary>
    BadUrl,
    /// <summary>A number. Its authored representation remains in the source.</summary>
    Number,
    /// <summary>A percentage.</summary>
    Percentage,
    /// <summary>A number followed by a unit identifier.</summary>
    Dimension,
    /// <summary>CSS whitespace.</summary>
    Whitespace,
    /// <summary>A comment retained as trivia; it is not whitespace.</summary>
    Comment,
    /// <summary>A colon.</summary>
    Colon,
    /// <summary>A semicolon.</summary>
    Semicolon,
    /// <summary>A comma.</summary>
    Comma,
    /// <summary>An opening parenthesis.</summary>
    OpenParenthesis,
    /// <summary>A closing parenthesis.</summary>
    CloseParenthesis,
    /// <summary>An opening square bracket.</summary>
    OpenBracket,
    /// <summary>A closing square bracket.</summary>
    CloseBracket,
    /// <summary>An opening curly brace.</summary>
    OpenBrace,
    /// <summary>A closing curly brace.</summary>
    CloseBrace,
    /// <summary>A CSS CDO token.</summary>
    Cdo,
    /// <summary>A CSS CDC token.</summary>
    Cdc,
    /// <summary>Any other delimiter.</summary>
    Delimiter,
    /// <summary>End of input, with a zero-length source span.</summary>
    EndOfFile
}

/// <summary>An immutable token with an exact span in the original UTF-16 source.</summary>
public readonly struct HtmlCssToken {
    internal HtmlCssToken(HtmlCssTokenKind kind, int offset, int length, string? value = null, bool identifierHash = false) {
        Kind = kind;
        Offset = offset;
        Length = length;
        Value = value;
        IsIdentifierHash = identifierHash;
    }

    /// <summary>The lexical category; tokenization alone does not imply property or rendering support.</summary>
    public HtmlCssTokenKind Kind { get; }
    /// <summary>Zero-based UTF-16 offset in the original input, before newline normalization.</summary>
    public int Offset { get; }
    /// <summary>Number of original UTF-16 characters, including delimiters and escapes.</summary>
    public int Length { get; }
    /// <summary>Decoded identifier, string, URL, dimension unit or delimiter; null for other categories.</summary>
    public string? Value { get; }
    /// <summary>Whether a hash token has the CSS identifier flag.</summary>
    public bool IsIdentifierHash { get; }
    /// <summary>Reads the exact authored token text from the source used to tokenize it.</summary>
    public string GetText(string source) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        return source.Substring(Offset, Length);
    }
}

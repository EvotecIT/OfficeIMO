using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Html.Css;

/// <summary>An exact UTF-16 range in the original CSS source.</summary>
public readonly struct HtmlCssSourceSpan : IEquatable<HtmlCssSourceSpan> {
    /// <summary>Creates a source span.</summary>
    public HtmlCssSourceSpan(int offset, int length) {
        if (offset < 0) throw new ArgumentOutOfRangeException(nameof(offset));
        if (length < 0) throw new ArgumentOutOfRangeException(nameof(length));
        Offset = offset;
        Length = length;
    }

    /// <summary>Zero-based UTF-16 offset.</summary>
    public int Offset { get; }
    /// <summary>Number of UTF-16 characters.</summary>
    public int Length { get; }
    /// <summary>Offset immediately after this span.</summary>
    public int End => checked(Offset + Length);
    /// <summary>Reads this span from the source it belongs to.</summary>
    public string GetText(string source) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (End > source.Length) throw new ArgumentOutOfRangeException(nameof(source));
        return source.Substring(Offset, Length);
    }
    /// <inheritdoc />
    public bool Equals(HtmlCssSourceSpan other) => Offset == other.Offset && Length == other.Length;
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is HtmlCssSourceSpan other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => unchecked(Offset * 397 ^ Length);
    /// <inheritdoc />
    public override string ToString() => $"{Offset}..{End}";
}

/// <summary>A one-based line and column together with the exact UTF-16 offset.</summary>
public readonly struct HtmlCssSourcePosition {
    internal HtmlCssSourcePosition(int offset, int line, int column) { Offset = offset; Line = line; Column = column; }
    /// <summary>Zero-based UTF-16 offset.</summary>
    public int Offset { get; }
    /// <summary>One-based source line.</summary>
    public int Line { get; }
    /// <summary>One-based UTF-16 source column.</summary>
    public int Column { get; }
    /// <inheritdoc />
    public override string ToString() => $"{Line}:{Column}";
}

/// <summary>Lossless CSS syntax categories. These do not imply selector, property, or rendering support.</summary>
public enum HtmlCssSyntaxKind {
    /// <summary>A token used as a component value.</summary>
    Token,
    /// <summary>A function and its nested component values.</summary>
    Function,
    /// <summary>A parenthesis, bracket, or brace block and its nested component values.</summary>
    SimpleBlock,
    /// <summary>An at-rule, whether known or unknown to a later consumer.</summary>
    AtRule,
    /// <summary>A qualified rule.</summary>
    QualifiedRule,
    /// <summary>A syntactically recognizable declaration.</summary>
    Declaration,
    /// <summary>Source retained after syntax recovery could not form a rule or declaration.</summary>
    Invalid
}

/// <summary>A diagnostic produced by CSS syntax recovery.</summary>
public sealed class HtmlCssSyntaxDiagnostic {
    internal HtmlCssSyntaxDiagnostic(string code, string message, HtmlCssSourceSpan span) { Code = code; Message = message; Span = span; }
    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }
    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }
    /// <summary>Source range associated with the diagnostic.</summary>
    public HtmlCssSourceSpan Span { get; }
}

/// <summary>Base type for a node in the lossless CSS syntax model.</summary>
public abstract class HtmlCssSyntaxNode {
    private readonly string _source;
    internal HtmlCssSyntaxNode(HtmlCssSyntaxKind kind, string source, HtmlCssSourceSpan span) {
        Kind = kind;
        _source = source;
        Span = span;
    }
    /// <summary>Syntax category.</summary>
    public HtmlCssSyntaxKind Kind { get; }
    /// <summary>Exact range in the original source.</summary>
    public HtmlCssSourceSpan Span { get; }
    /// <summary>Original source retained by derived source-slice views.</summary>
    protected string Source => _source;
    /// <summary>Returns the exact authored text represented by this node.</summary>
    public string GetText() => Span.GetText(_source);
    /// <inheritdoc />
    public override string ToString() => GetText();
}

/// <summary>Base type for a token, function, or nested simple block used in a prelude or value.</summary>
public abstract class HtmlCssComponentValue : HtmlCssSyntaxNode {
    internal HtmlCssComponentValue(HtmlCssSyntaxKind kind, string source, HtmlCssSourceSpan span) : base(kind, source, span) { }
}

/// <summary>A lexical token used as a component value.</summary>
public sealed class HtmlCssTokenValue : HtmlCssComponentValue {
    internal HtmlCssTokenValue(string source, HtmlCssToken token) : base(HtmlCssSyntaxKind.Token, source, new HtmlCssSourceSpan(token.Offset, token.Length)) { Token = token; }
    /// <summary>The retained lexical token.</summary>
    public HtmlCssToken Token { get; }
}

/// <summary>A function with recursively grouped component values.</summary>
public sealed class HtmlCssFunctionValue : HtmlCssComponentValue {
    internal HtmlCssFunctionValue(string source, HtmlCssSourceSpan span, string name, IReadOnlyList<HtmlCssComponentValue> values, bool isClosed)
        : base(HtmlCssSyntaxKind.Function, source, span) { Name = name; Values = values; IsClosed = isClosed; }
    /// <summary>Decoded function name.</summary>
    public string Name { get; }
    /// <summary>Ordered component values inside the parentheses.</summary>
    public IReadOnlyList<HtmlCssComponentValue> Values { get; }
    /// <summary>Whether a closing parenthesis occurred before the surrounding boundary or end of input.</summary>
    public bool IsClosed { get; }
}

/// <summary>A recursively grouped parenthesis, bracket, or brace block.</summary>
public sealed class HtmlCssSimpleBlock : HtmlCssComponentValue {
    internal HtmlCssSimpleBlock(
        string source,
        HtmlCssSourceSpan span,
        HtmlCssTokenKind openingKind,
        HtmlCssTokenKind closingKind,
        IReadOnlyList<HtmlCssComponentValue> values,
        bool isClosed)
        : base(HtmlCssSyntaxKind.SimpleBlock, source, span) {
        OpeningKind = openingKind;
        ClosingKind = closingKind;
        Values = values;
        IsClosed = isClosed;
    }
    /// <summary>Opening delimiter token kind.</summary>
    public HtmlCssTokenKind OpeningKind { get; }
    /// <summary>Expected closing delimiter token kind.</summary>
    public HtmlCssTokenKind ClosingKind { get; }
    /// <summary>Ordered nested component values.</summary>
    public IReadOnlyList<HtmlCssComponentValue> Values { get; }
    /// <summary>Whether the expected closing delimiter was present.</summary>
    public bool IsClosed { get; }
}

internal static class HtmlCssSyntaxCollections {
    internal static IReadOnlyList<T> ReadOnly<T>(List<T> values) => new ReadOnlyCollection<T>(values);
}

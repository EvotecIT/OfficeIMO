using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.Html.Css;

/// <summary>Base type for a source-preserving CSS syntax result.</summary>
public abstract class HtmlCssSyntaxDocument {
    private readonly int[] _lineStarts;
    internal HtmlCssSyntaxDocument(
        string source,
        IReadOnlyList<HtmlCssToken> tokens,
        IReadOnlyList<HtmlCssSyntaxNode> contents,
        IReadOnlyList<HtmlCssSyntaxDiagnostic> diagnostics) {
        Source = source;
        Tokens = tokens;
        Contents = contents;
        Diagnostics = diagnostics;
        var starts = new List<int> { 0 };
        for (int index = 0; index < source.Length; index++) {
            if (source[index] == '\r') {
                if (index + 1 < source.Length && source[index + 1] == '\n') index++;
                starts.Add(index + 1);
            } else if (source[index] == '\n' || source[index] == '\f') starts.Add(index + 1);
        }
        _lineStarts = starts.ToArray();
    }

    /// <summary>Exact original CSS input.</summary>
    public string Source { get; }
    /// <summary>All lexical tokens, including trivia and the final end-of-input token.</summary>
    public IReadOnlyList<HtmlCssToken> Tokens { get; }
    /// <summary>Rules, declarations, and invalid recovered source in authored order.</summary>
    public IReadOnlyList<HtmlCssSyntaxNode> Contents { get; }
    /// <summary>Syntax recovery diagnostics. Unknown names and values do not produce diagnostics.</summary>
    public IReadOnlyList<HtmlCssSyntaxDiagnostic> Diagnostics { get; }
    /// <summary>Maps an exact UTF-16 offset to a one-based line and column.</summary>
    public HtmlCssSourcePosition GetPosition(int offset) {
        if (offset < 0 || offset > Source.Length) throw new ArgumentOutOfRangeException(nameof(offset));
        int line = Array.BinarySearch(_lineStarts, offset);
        if (line < 0) line = ~line - 1;
        return new HtmlCssSourcePosition(offset, line + 1, offset - _lineStarts[line] + 1);
    }
    /// <summary>Returns the original input unchanged.</summary>
    public string ToCss() => Source;
    /// <inheritdoc />
    public override string ToString() => Source;
}

/// <summary>A source-preserving CSS stylesheet syntax result.</summary>
public sealed class HtmlCssStyleSheet : HtmlCssSyntaxDocument {
    internal HtmlCssStyleSheet(
        string source,
        IReadOnlyList<HtmlCssToken> tokens,
        IReadOnlyList<HtmlCssSyntaxNode> contents,
        IReadOnlyList<HtmlCssSyntaxDiagnostic> diagnostics)
        : base(source, tokens, contents, diagnostics) {
        Rules = new ReadOnlyCollection<HtmlCssRule>(contents.OfType<HtmlCssRule>().ToList());
    }
    /// <summary>Top-level at-rules and qualified rules.</summary>
    public IReadOnlyList<HtmlCssRule> Rules { get; }
}

/// <summary>A source-preserving CSS style-block result for inline declarations or nested rule content.</summary>
public sealed class HtmlCssStyleBlock : HtmlCssSyntaxDocument {
    internal HtmlCssStyleBlock(
        string source,
        IReadOnlyList<HtmlCssToken> tokens,
        IReadOnlyList<HtmlCssSyntaxNode> contents,
        IReadOnlyList<HtmlCssSyntaxDiagnostic> diagnostics)
        : base(source, tokens, contents, diagnostics) {
        Declarations = new ReadOnlyCollection<HtmlCssDeclaration>(contents.OfType<HtmlCssDeclaration>().ToList());
    }
    /// <summary>Direct declarations in authored order. Nested rules remain in <see cref="HtmlCssSyntaxDocument.Contents"/>.</summary>
    public IReadOnlyList<HtmlCssDeclaration> Declarations { get; }
}

/// <summary>Resource bounds for an owned CSS syntax parse.</summary>
public sealed class HtmlCssSyntaxOptions {
    /// <summary>Maximum UTF-16 input length, or null for caller-bounded input.</summary>
    public int? MaxInputCharacters { get; set; } = 8 * 1024 * 1024;
    /// <summary>Maximum lexical tokens excluding end-of-input, or null to disable the token limit.</summary>
    public int? MaxTokens { get; set; } = 1_000_000;
    /// <summary>Maximum nested function/block depth, or null to disable the nesting limit.</summary>
    public int? MaxNestingDepth { get; set; } = 256;
    /// <summary>Maximum materialized syntax nodes, or null to disable the syntax-node limit.</summary>
    public int? MaxSyntaxNodes { get; set; } = 1_000_000;
    /// <summary>Copies these settings.</summary>
    public HtmlCssSyntaxOptions Clone() => new HtmlCssSyntaxOptions {
        MaxInputCharacters = MaxInputCharacters,
        MaxTokens = MaxTokens,
        MaxNestingDepth = MaxNestingDepth,
        MaxSyntaxNodes = MaxSyntaxNodes
    };
    /// <summary>Rejects nonpositive limits.</summary>
    public void Validate() {
        if (MaxInputCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxInputCharacters));
        if (MaxTokens <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTokens));
        if (MaxNestingDepth <= 0) throw new ArgumentOutOfRangeException(nameof(MaxNestingDepth));
        if (MaxSyntaxNodes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSyntaxNodes));
    }
}

/// <summary>A CSS syntax operation exceeded a requested resource bound.</summary>
public sealed class HtmlCssSyntaxLimitException : InvalidOperationException {
    /// <summary>Creates a failure describing the exhausted budget.</summary>
    public HtmlCssSyntaxLimitException(string limitName, long actual, long maximum)
        : base($"CSS syntax {limitName} limit exceeded ({actual} > {maximum}).") { LimitName = limitName; Actual = actual; Maximum = maximum; }
    /// <summary>The exhausted option.</summary>
    public string LimitName { get; }
    /// <summary>The observed amount of work.</summary>
    public long Actual { get; }
    /// <summary>The configured ceiling.</summary>
    public long Maximum { get; }
}

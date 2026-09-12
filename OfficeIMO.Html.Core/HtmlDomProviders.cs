using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Html.Dom;

/// <summary>Parses inert HTML into an owned document. Implementations must perform no resource loading or script execution.</summary>
public interface IHtmlParserProvider {
    /// <summary>Identifies the parser implementation for diagnostics and reproducibility.</summary>
    string Id { get; }
    /// <summary>Parses a complete document and returns an immutable snapshot.</summary>
    HtmlDocument Parse(string source, HtmlParseOptions options, CancellationToken cancellationToken = default);
}

/// <summary>Replaceable syntax services used by an owned document without exposing provider nodes.</summary>
public interface IHtmlDomServices {
    /// <summary>Serializes a node or its children as HTML without executing or loading content.</summary>
    string Serialize(HtmlNode node, bool childrenOnly = false);
    /// <summary>Returns matching descendant elements in document order, using the supplied scope.</summary>
    IReadOnlyList<HtmlElement> QuerySelectorAll(HtmlNode scope, string selector);
    /// <summary>Tests one element against a selector.</summary>
    bool Matches(HtmlElement element, string selector);
}

/// <summary>Parser-local source and tree budgets. Conversion callers translate their shared budgets into this contract.</summary>
public sealed class HtmlParseOptions {
    /// <summary>Maximum UTF-16 input characters, or null for no source-length limit.</summary>
    public int? MaxInputCharacters { get; set; } = 72 * 1024 * 1024;
    /// <summary>Maximum tree nodes, excluding the document itself, or null for no node limit.</summary>
    public int? MaxNodes { get; set; } = 100_000;
    /// <summary>Maximum element nesting depth, with the root element at depth one, or null for no depth limit.</summary>
    public int? MaxDepth { get; set; } = 256;

    /// <summary>Checks that configured limits are positive.</summary>
    public void Validate() {
        if (MaxInputCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxInputCharacters));
        if (MaxNodes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxNodes));
        if (MaxDepth <= 0) throw new ArgumentOutOfRangeException(nameof(MaxDepth));
    }

    /// <summary>Copies the settings so a provider can retain a stable request.</summary>
    public HtmlParseOptions Clone() => new HtmlParseOptions { MaxInputCharacters = MaxInputCharacters, MaxNodes = MaxNodes, MaxDepth = MaxDepth };
}

/// <summary>A parser could not complete within a caller-specified source or tree budget.</summary>
public sealed class HtmlParseLimitException : InvalidOperationException {
    /// <summary>Creates a limit failure with observed and configured values.</summary>
    public HtmlParseLimitException(string limitName, long actual, long maximum) : base($"HTML {limitName} limit exceeded ({actual} > {maximum}).") {
        LimitName = limitName;
        Actual = actual;
        Maximum = maximum;
    }
    /// <summary>The exhausted parser option.</summary>
    public string LimitName { get; }
    /// <summary>The observed amount of work or input.</summary>
    public long Actual { get; }
    /// <summary>The configured maximum.</summary>
    public long Maximum { get; }
}

using System.Collections.Generic;

namespace OfficeIMO.Html.Css;

/// <summary>Base type for a CSS at-rule or qualified rule.</summary>
public abstract class HtmlCssRule : HtmlCssSyntaxNode {
    internal HtmlCssRule(
        HtmlCssSyntaxKind kind,
        string source,
        HtmlCssSourceSpan span,
        IReadOnlyList<HtmlCssComponentValue> prelude,
        HtmlCssSimpleBlock? block,
        IReadOnlyList<HtmlCssSyntaxNode> contents)
        : base(kind, source, span) { Prelude = prelude; Block = block; Contents = contents; }
    /// <summary>Component values before the rule block or terminator.</summary>
    public IReadOnlyList<HtmlCssComponentValue> Prelude { get; }
    /// <summary>The exact curly block, or null for a statement at-rule.</summary>
    public HtmlCssSimpleBlock? Block { get; }
    /// <summary>Best-effort syntax-only declarations and nested rules in a block, in authored order.</summary>
    /// <remarks>The raw <see cref="Block"/> remains authoritative for at-rules whose grammar is not known here.</remarks>
    public IReadOnlyList<HtmlCssSyntaxNode> Contents { get; }
}

/// <summary>An at-rule retained without requiring the name or grammar to be recognized.</summary>
public sealed class HtmlCssAtRule : HtmlCssRule {
    internal HtmlCssAtRule(
        string source,
        HtmlCssSourceSpan span,
        string name,
        IReadOnlyList<HtmlCssComponentValue> prelude,
        HtmlCssSimpleBlock? block,
        IReadOnlyList<HtmlCssSyntaxNode> contents)
        : base(HtmlCssSyntaxKind.AtRule, source, span, prelude, block, contents) { Name = name; }
    /// <summary>Decoded at-rule name.</summary>
    public string Name { get; }
    /// <summary>Whether the rule ended with a semicolon or end of input instead of a block.</summary>
    public bool IsStatement => Block == null;
}

/// <summary>A qualified rule whose prelude and block are retained without selector validation.</summary>
public sealed class HtmlCssQualifiedRule : HtmlCssRule {
    internal HtmlCssQualifiedRule(
        string source,
        HtmlCssSourceSpan span,
        IReadOnlyList<HtmlCssComponentValue> prelude,
        HtmlCssSimpleBlock block,
        IReadOnlyList<HtmlCssSyntaxNode> contents)
        : base(HtmlCssSyntaxKind.QualifiedRule, source, span, prelude, block, contents) { }
}

/// <summary>A declaration retained in source order without applying property grammar.</summary>
public sealed class HtmlCssDeclaration : HtmlCssSyntaxNode {
    private string? _valueText;

    internal HtmlCssDeclaration(
        string source,
        HtmlCssSourceSpan span,
        string name,
        HtmlCssSourceSpan valueSpan,
        IReadOnlyList<HtmlCssComponentValue> values,
        bool isImportant)
        : base(HtmlCssSyntaxKind.Declaration, source, span) {
        Name = name;
        ValueSpan = valueSpan;
        Values = values;
        IsImportant = isImportant;
    }
    /// <summary>Decoded declaration name. Custom-property names remain case-sensitive.</summary>
    public string Name { get; }
    /// <summary>Exact source range after the colon and before the declaration terminator.</summary>
    public HtmlCssSourceSpan ValueSpan { get; }
    /// <summary>Exact authored text after the colon and before the declaration terminator.</summary>
    public string ValueText => _valueText ??= ValueSpan.GetText(Source);
    /// <summary>Nested component values in authored order, including comments and whitespace.</summary>
    public IReadOnlyList<HtmlCssComponentValue> Values { get; }
    /// <summary>Whether the final non-trivia value tokens form <c>!important</c>.</summary>
    public bool IsImportant { get; }
    /// <summary>Whether this is a case-sensitive custom property.</summary>
    public bool IsCustomProperty => Name.StartsWith("--", System.StringComparison.Ordinal);
}

/// <summary>Exact source retained after error recovery could not form a declaration or rule.</summary>
public sealed class HtmlCssInvalidSyntax : HtmlCssSyntaxNode {
    internal HtmlCssInvalidSyntax(string source, HtmlCssSourceSpan span, string diagnosticCode)
        : base(HtmlCssSyntaxKind.Invalid, source, span) { DiagnosticCode = diagnosticCode; }
    /// <summary>Code of the recovery diagnostic that produced this node.</summary>
    public string DiagnosticCode { get; }
}

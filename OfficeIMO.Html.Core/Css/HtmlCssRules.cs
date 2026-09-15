using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.Html.Css;

/// <summary>Base type for a CSS at-rule or qualified rule.</summary>
public abstract class HtmlCssRule : HtmlCssSyntaxNode {
    private IReadOnlyList<HtmlCssDeclaration>? _declarations;
    private string? _preludeText;
    private IReadOnlyList<HtmlCssRule>? _rules;
    internal HtmlCssRule(
        HtmlCssSyntaxKind kind,
        string source,
        HtmlCssSourceSpan span,
        IReadOnlyList<HtmlCssComponentValue> prelude,
        HtmlCssSimpleBlock? block,
        IReadOnlyList<HtmlCssSyntaxNode> contents)
        : base(kind, source, span) {
        Prelude = prelude;
        Block = block;
        Contents = contents;
    }
    /// <summary>Component values before the rule block or terminator.</summary>
    public IReadOnlyList<HtmlCssComponentValue> Prelude { get; }
    /// <summary>Exact source text before the rule block or at-rule terminator.</summary>
    public string PreludeText {
        get {
            if (_preludeText != null) return _preludeText;
            var text = new System.Text.StringBuilder();
            foreach (HtmlCssComponentValue value in Prelude) text.Append(value.GetText());
            return _preludeText = text.ToString();
        }
    }
    /// <summary>The exact curly block, or null for a statement at-rule.</summary>
    public HtmlCssSimpleBlock? Block { get; }
    /// <summary>Best-effort syntax-only declarations and nested rules in a block, in authored order.</summary>
    /// <remarks>The raw <see cref="Block"/> remains authoritative for at-rules whose grammar is not known here.</remarks>
    public IReadOnlyList<HtmlCssSyntaxNode> Contents { get; }
    /// <summary>Direct declarations in the rule block, in authored order.</summary>
    /// <remarks>Declarations inside nested rules are exposed by those rules. Unknown at-rule grammar remains best-effort.</remarks>
    public IReadOnlyList<HtmlCssDeclaration> Declarations =>
        _declarations ??= new ReadOnlyCollection<HtmlCssDeclaration>(Contents.OfType<HtmlCssDeclaration>().ToList());
    /// <summary>Direct nested at-rules and qualified rules, in authored order.</summary>
    /// <remarks>The raw <see cref="Block"/> remains authoritative when an unknown at-rule uses a different block grammar.</remarks>
    public IReadOnlyList<HtmlCssRule> Rules =>
        _rules ??= new ReadOnlyCollection<HtmlCssRule>(Contents.OfType<HtmlCssRule>().ToList());
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

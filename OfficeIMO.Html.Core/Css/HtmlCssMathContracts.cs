using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Html.Css;

/// <summary>Length units accepted by the owned static-rendering subset.</summary>
public enum HtmlCssLengthUnit {
    /// <summary>CSS pixels.</summary>
    Px,
    /// <summary>Points.</summary>
    Pt,
    /// <summary>Picas.</summary>
    Pc,
    /// <summary>Inches.</summary>
    In,
    /// <summary>Centimeters.</summary>
    Cm,
    /// <summary>Millimeters.</summary>
    Mm,
    /// <summary>Quarter millimeters.</summary>
    Q,
    /// <summary>The element font size.</summary>
    Em,
    /// <summary>The root element font size.</summary>
    Rem,
    /// <summary>One percent of the default viewport width.</summary>
    Vw,
    /// <summary>One percent of the default viewport height.</summary>
    Vh,
    /// <summary>One percent of the smaller default viewport dimension.</summary>
    Vmin,
    /// <summary>One percent of the larger default viewport dimension.</summary>
    Vmax,
    /// <summary>One percent of the small viewport width.</summary>
    Svw,
    /// <summary>One percent of the small viewport height.</summary>
    Svh,
    /// <summary>One percent of the smaller small-viewport dimension.</summary>
    Svmin,
    /// <summary>One percent of the larger small-viewport dimension.</summary>
    Svmax,
    /// <summary>One percent of the large viewport width.</summary>
    Lvw,
    /// <summary>One percent of the large viewport height.</summary>
    Lvh,
    /// <summary>One percent of the smaller large-viewport dimension.</summary>
    Lvmin,
    /// <summary>One percent of the larger large-viewport dimension.</summary>
    Lvmax,
    /// <summary>One percent of the dynamic viewport width.</summary>
    Dvw,
    /// <summary>One percent of the dynamic viewport height.</summary>
    Dvh,
    /// <summary>One percent of the smaller dynamic-viewport dimension.</summary>
    Dvmin,
    /// <summary>One percent of the larger dynamic-viewport dimension.</summary>
    Dvmax,
    /// <summary>One percent of the query container width.</summary>
    Cqw,
    /// <summary>One percent of the query container height.</summary>
    Cqh,
    /// <summary>One percent of the query container inline size.</summary>
    Cqi,
    /// <summary>One percent of the query container block size.</summary>
    Cqb,
    /// <summary>One percent of the smaller query-container dimension.</summary>
    Cqmin,
    /// <summary>One percent of the larger query-container dimension.</summary>
    Cqmax
}

/// <summary>Syntax node kinds in a parsed CSS numeric expression.</summary>
public enum HtmlCssMathExpressionKind {
    /// <summary>A number, percentage, or dimension literal.</summary>
    Literal,
    /// <summary>An addition.</summary>
    Add,
    /// <summary>A subtraction.</summary>
    Subtract,
    /// <summary>A multiplication.</summary>
    Multiply,
    /// <summary>A division.</summary>
    Divide,
    /// <summary>A calc() grouping.</summary>
    Calc,
    /// <summary>A min() comparison.</summary>
    Min,
    /// <summary>A max() comparison.</summary>
    Max,
    /// <summary>A clamp() comparison.</summary>
    Clamp
}

/// <summary>An immutable, provider-independent CSS numeric expression.</summary>
public sealed class HtmlCssMathExpression {
    internal HtmlCssMathExpression(
        HtmlCssMathExpressionKind kind,
        HtmlCssNumericType type,
        string canonicalText,
        double? value = null,
        HtmlCssLengthUnit? unit = null,
        IReadOnlyList<HtmlCssMathExpression>? children = null) {
        Kind = kind;
        Type = type;
        CanonicalText = canonicalText;
        Value = value;
        Unit = unit;
        Children = children ?? EmptyChildren;
    }

    private static readonly IReadOnlyList<HtmlCssMathExpression> EmptyChildren =
        new ReadOnlyCollection<HtmlCssMathExpression>(new List<HtmlCssMathExpression>());

    /// <summary>The expression node kind.</summary>
    public HtmlCssMathExpressionKind Kind { get; }
    /// <summary>The numeric type established without resolving layout-dependent context.</summary>
    public HtmlCssNumericType Type { get; }
    /// <summary>A stable invariant representation of this implemented expression subset.</summary>
    public string CanonicalText { get; }
    /// <summary>The finite literal value, or null for a compound expression.</summary>
    public double? Value { get; }
    /// <summary>The unit of a length literal, or null for numbers and percentages.</summary>
    public HtmlCssLengthUnit? Unit { get; }
    /// <summary>Child expressions in source order.</summary>
    public IReadOnlyList<HtmlCssMathExpression> Children { get; }
    /// <summary>Whether the expression contains a CSS math function or arithmetic operator.</summary>
    public bool IsCalculated => Kind != HtmlCssMathExpressionKind.Literal;
}

/// <summary>Outcome of parsing a standalone CSS numeric expression.</summary>
public enum HtmlCssMathParseStatus {
    /// <summary>The expression is in the implemented grammar and type subset.</summary>
    Parsed,
    /// <summary>The expression is malformed.</summary>
    InvalidSyntax,
    /// <summary>The expression uses a numeric unit or function outside the implemented subset.</summary>
    UnsupportedValue,
    /// <summary>The expression combines incompatible numeric types.</summary>
    IncompatibleTypes
}

/// <summary>Result of parsing a standalone CSS numeric expression.</summary>
public sealed class HtmlCssMathParseResult {
    internal HtmlCssMathParseResult(string authoredText, HtmlCssMathParseStatus status, HtmlCssMathExpression? expression) {
        AuthoredText = authoredText;
        Status = status;
        Expression = expression;
    }

    /// <summary>Trimmed source supplied by the caller.</summary>
    public string AuthoredText { get; }
    /// <summary>Parse outcome.</summary>
    public HtmlCssMathParseStatus Status { get; }
    /// <summary>The complete expression when parsing succeeded.</summary>
    public HtmlCssMathExpression? Expression { get; }
    /// <summary>Whether a complete typed expression was produced.</summary>
    public bool IsParsed => Status == HtmlCssMathParseStatus.Parsed;
}

/// <summary>Resource limits for one standalone CSS math parse.</summary>
public sealed class HtmlCssMathOptions {
    /// <summary>Maximum UTF-16 input length.</summary>
    public int MaxInputCharacters { get; set; } = 4096;
    /// <summary>Maximum lexical tokens, excluding end-of-input.</summary>
    public int MaxTokens { get; set; } = 512;
    /// <summary>Maximum nested parentheses and math functions.</summary>
    public int MaxNestingDepth { get; set; } = 32;
    /// <summary>Maximum arithmetic, function, and argument-separator operations.</summary>
    public int MaxOperations { get; set; } = 256;
    /// <summary>Maximum arguments accepted by min() or max().</summary>
    public int MaxArguments { get; set; } = 32;

    /// <summary>Copies the request settings.</summary>
    public HtmlCssMathOptions Clone() => new HtmlCssMathOptions {
        MaxInputCharacters = MaxInputCharacters,
        MaxTokens = MaxTokens,
        MaxNestingDepth = MaxNestingDepth,
        MaxOperations = MaxOperations,
        MaxArguments = MaxArguments
    };

    internal void Validate() {
        if (MaxInputCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxInputCharacters));
        if (MaxTokens <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTokens));
        if (MaxNestingDepth <= 0) throw new ArgumentOutOfRangeException(nameof(MaxNestingDepth));
        if (MaxOperations <= 0) throw new ArgumentOutOfRangeException(nameof(MaxOperations));
        if (MaxArguments <= 0) throw new ArgumentOutOfRangeException(nameof(MaxArguments));
    }
}

/// <summary>A CSS math parse exceeded an explicit resource bound.</summary>
public sealed class HtmlCssMathLimitException : InvalidOperationException {
    /// <summary>Creates a failure describing the exhausted budget.</summary>
    public HtmlCssMathLimitException(string limitName, long actual, long maximum)
        : base($"CSS math {limitName} limit exceeded ({actual} > {maximum}).") {
        LimitName = limitName;
        Actual = actual;
        Maximum = maximum;
    }

    /// <summary>The exhausted option.</summary>
    public string LimitName { get; }
    /// <summary>The observed value.</summary>
    public long Actual { get; }
    /// <summary>The configured ceiling.</summary>
    public long Maximum { get; }
}

/// <summary>Context used to resolve owned length and percentage expressions to CSS pixels.</summary>
public sealed class HtmlCssLengthResolutionContext {
    /// <summary>Length represented by 100% for the consuming property.</summary>
    public double? PercentageReference { get; set; }
    /// <summary>Computed element font size in CSS pixels.</summary>
    public double? FontSize { get; set; }
    /// <summary>Computed root font size in CSS pixels.</summary>
    public double? RootFontSize { get; set; }
    /// <summary>Default viewport width in CSS pixels. Specialized viewport sizes fall back to this value.</summary>
    public double? ViewportWidth { get; set; }
    /// <summary>Default viewport height in CSS pixels. Specialized viewport sizes fall back to this value.</summary>
    public double? ViewportHeight { get; set; }
    /// <summary>Small viewport width in CSS pixels.</summary>
    public double? SmallViewportWidth { get; set; }
    /// <summary>Small viewport height in CSS pixels.</summary>
    public double? SmallViewportHeight { get; set; }
    /// <summary>Large viewport width in CSS pixels.</summary>
    public double? LargeViewportWidth { get; set; }
    /// <summary>Large viewport height in CSS pixels.</summary>
    public double? LargeViewportHeight { get; set; }
    /// <summary>Dynamic viewport width in CSS pixels.</summary>
    public double? DynamicViewportWidth { get; set; }
    /// <summary>Dynamic viewport height in CSS pixels.</summary>
    public double? DynamicViewportHeight { get; set; }
    /// <summary>Query-container width in CSS pixels.</summary>
    public double? ContainerWidth { get; set; }
    /// <summary>Query-container height in CSS pixels.</summary>
    public double? ContainerHeight { get; set; }
    /// <summary>Query-container inline size in CSS pixels. Falls back to container width.</summary>
    public double? ContainerInlineSize { get; set; }
    /// <summary>Query-container block size in CSS pixels. Falls back to container height.</summary>
    public double? ContainerBlockSize { get; set; }
}

/// <summary>Outcome of resolving an owned CSS length expression.</summary>
public enum HtmlCssLengthResolutionStatus {
    /// <summary>The expression resolved to a finite CSS-pixel value.</summary>
    Resolved,
    /// <summary>The consuming property's percentage reference is unavailable.</summary>
    MissingPercentageReference,
    /// <summary>The element font size is unavailable.</summary>
    MissingFontSize,
    /// <summary>The root font size is unavailable.</summary>
    MissingRootFontSize,
    /// <summary>A required viewport width is unavailable.</summary>
    MissingViewportWidth,
    /// <summary>A required viewport height is unavailable.</summary>
    MissingViewportHeight,
    /// <summary>A supplied context value or calculated result is not finite.</summary>
    NonFiniteValue,
    /// <summary>The expression does not resolve to a length in this API.</summary>
    IncompatibleType
}

/// <summary>Finite resolution result or an explicit reason why context was insufficient.</summary>
public sealed class HtmlCssLengthResolutionResult {
    internal HtmlCssLengthResolutionResult(HtmlCssLengthResolutionStatus status, double? value) {
        Status = status;
        Value = value;
    }

    /// <summary>Resolution outcome.</summary>
    public HtmlCssLengthResolutionStatus Status { get; }
    /// <summary>Resolved CSS pixels, or null when resolution did not complete.</summary>
    public double? Value { get; }
    /// <summary>Whether a finite CSS-pixel value is available.</summary>
    public bool IsResolved => Status == HtmlCssLengthResolutionStatus.Resolved;
}

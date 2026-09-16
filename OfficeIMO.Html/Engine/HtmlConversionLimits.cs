namespace OfficeIMO.Html;

/// <summary>
/// Shared safety and complexity limits applied before HTML is analyzed by target adapters.
/// </summary>
/// <remarks>
/// This is the single owner for source, DOM, stylesheet, selector, responsive-resource, and
/// semantic-metadata budgets. Target adapters may add native-format limits, but should not
/// independently reinterpret these shared HTML boundaries.
/// </remarks>
public sealed class HtmlConversionLimits {
    internal const int DefaultMaxResponsiveImageCandidates = 64;
    internal const int DefaultMaxResponsiveImageSizesCharacters = 64 * 1024;

    /// <summary>Creates conservative limits suitable for untrusted HTML ingestion.</summary>
    public static HtmlConversionLimits CreateUntrustedProfile() => new HtmlConversionLimits {
        MaxInputCharacters = 72 * 1024 * 1024,
        MaxHtmlNodes = 100_000,
        MaxHtmlDepth = 256,
        // Large embedded font data URIs are common in offline HTML/PDF workflows. Complexity
        // budgets below still bound parsed rules, declarations, and selector work.
        MaxCssBytes = 72L * 1024L * 1024L,
        MaxTotalCssBytes = 72L * 1024L * 1024L,
        MaxCssRules = 10_000,
        MaxCssDeclarations = 100_000,
        MaxCssTokens = 1_000_000,
        MaxCssNestingDepth = 64,
        MaxCssSyntaxNodes = 1_000_000,
        MaxCssSelectorCharacters = 64 * 1024,
        MaxCssSelectorsPerRule = 256,
        MaxSelectorEvaluations = 10_000_000L,
        MaxResponsiveImageCandidates = DefaultMaxResponsiveImageCandidates,
        MaxResponsiveImageSizesCharacters = DefaultMaxResponsiveImageSizesCharacters,
        MaxSemanticMetadataCharacters = 1024 * 1024
    };

    /// <summary>Creates compatibility-oriented limits for caller-trusted HTML.</summary>
    public static HtmlConversionLimits CreateTrustedProfile() => new HtmlConversionLimits();

    /// <summary>Maximum UTF-16 source characters, or <c>null</c> for no source-length limit.</summary>
    /// <remarks>Owned conversion input also checks aggregate attached node names and values before capture,
    /// then bounds canonical HTML as it is serialized. Namespace URIs are not counted as node data.</remarks>
    public int? MaxInputCharacters { get; set; }

    /// <summary>Maximum DOM nodes, or <c>null</c> for no node limit.</summary>
    public int? MaxHtmlNodes { get; set; }

    /// <summary>Maximum DOM nesting depth, or <c>null</c> for no depth limit.</summary>
    public int? MaxHtmlDepth { get; set; }

    /// <summary>Maximum UTF-8 bytes in one embedded stylesheet, or <c>null</c> for no per-sheet limit.</summary>
    public long? MaxCssBytes { get; set; }

    /// <summary>Maximum UTF-8 bytes across embedded stylesheets, or <c>null</c> for no total limit.</summary>
    public long? MaxTotalCssBytes { get; set; }

    /// <summary>Maximum active CSS rules, or <c>null</c> for no rule-count limit.</summary>
    public int? MaxCssRules { get; set; }

    /// <summary>Maximum declarations across active CSS rules, or <c>null</c> for no declaration limit.</summary>
    public int? MaxCssDeclarations { get; set; }

    /// <summary>Maximum lexical tokens produced while parsing one inline declaration block, or <c>null</c> for no token limit.</summary>
    public int? MaxCssTokens { get; set; }

    /// <summary>Maximum nested CSS grouping/rule-block depth. Defaults to 256 even for trusted input.</summary>
    public int? MaxCssNestingDepth { get; set; } = 256;

    /// <summary>Maximum syntax nodes materialized while parsing one inline declaration block, or <c>null</c> for no node limit.</summary>
    public int? MaxCssSyntaxNodes { get; set; }

    /// <summary>Maximum UTF-16 characters in one selector list after CSS nesting is resolved.</summary>
    public int? MaxCssSelectorCharacters { get; set; } = 64 * 1024;

    /// <summary>Maximum selectors in one rule after CSS nesting is resolved.</summary>
    public int? MaxCssSelectorsPerRule { get; set; } = 256;

    /// <summary>Maximum element/selector match attempts, or <c>null</c> for no evaluation limit.</summary>
    public long? MaxSelectorEvaluations { get; set; }

    /// <summary>Maximum responsive image candidates per source set, or <c>null</c> for no candidate limit.</summary>
    public int? MaxResponsiveImageCandidates { get; set; }

    /// <summary>Maximum UTF-16 characters parsed from one responsive image <c>sizes</c> value, or <c>null</c> for no limit.</summary>
    public int? MaxResponsiveImageSizesCharacters { get; set; }

    /// <summary>Maximum characters accepted from one semantic metadata field.</summary>
    public int? MaxSemanticMetadataCharacters { get; set; }

    /// <summary>Creates an independent limits snapshot.</summary>
    public HtmlConversionLimits Clone() => new HtmlConversionLimits {
        MaxInputCharacters = MaxInputCharacters,
        MaxHtmlNodes = MaxHtmlNodes,
        MaxHtmlDepth = MaxHtmlDepth,
        MaxCssBytes = MaxCssBytes,
        MaxTotalCssBytes = MaxTotalCssBytes,
        MaxCssRules = MaxCssRules,
        MaxCssDeclarations = MaxCssDeclarations,
        MaxCssTokens = MaxCssTokens,
        MaxCssNestingDepth = MaxCssNestingDepth,
        MaxCssSyntaxNodes = MaxCssSyntaxNodes,
        MaxCssSelectorCharacters = MaxCssSelectorCharacters,
        MaxCssSelectorsPerRule = MaxCssSelectorsPerRule,
        MaxSelectorEvaluations = MaxSelectorEvaluations,
        MaxResponsiveImageCandidates = MaxResponsiveImageCandidates,
        MaxResponsiveImageSizesCharacters = MaxResponsiveImageSizesCharacters,
        MaxSemanticMetadataCharacters = MaxSemanticMetadataCharacters
    };

    /// <summary>
    /// Combines two shared limit sets without allowing either boundary to be relaxed.
    /// </summary>
    internal static HtmlConversionLimits Intersect(HtmlConversionLimits? first, HtmlConversionLimits? second) {
        HtmlConversionLimits left = first ?? CreateTrustedProfile();
        HtmlConversionLimits right = second ?? CreateTrustedProfile();
        return new HtmlConversionLimits {
            MaxInputCharacters = Minimum(left.MaxInputCharacters, right.MaxInputCharacters),
            MaxHtmlNodes = Minimum(left.MaxHtmlNodes, right.MaxHtmlNodes),
            MaxHtmlDepth = Minimum(left.MaxHtmlDepth, right.MaxHtmlDepth),
            MaxCssBytes = Minimum(left.MaxCssBytes, right.MaxCssBytes),
            MaxTotalCssBytes = Minimum(left.MaxTotalCssBytes, right.MaxTotalCssBytes),
            MaxCssRules = Minimum(left.MaxCssRules, right.MaxCssRules),
            MaxCssDeclarations = Minimum(left.MaxCssDeclarations, right.MaxCssDeclarations),
            MaxCssTokens = Minimum(left.MaxCssTokens, right.MaxCssTokens),
            MaxCssNestingDepth = Minimum(left.MaxCssNestingDepth, right.MaxCssNestingDepth),
            MaxCssSyntaxNodes = Minimum(left.MaxCssSyntaxNodes, right.MaxCssSyntaxNodes),
            MaxCssSelectorCharacters = Minimum(left.MaxCssSelectorCharacters, right.MaxCssSelectorCharacters),
            MaxCssSelectorsPerRule = Minimum(left.MaxCssSelectorsPerRule, right.MaxCssSelectorsPerRule),
            MaxSelectorEvaluations = Minimum(left.MaxSelectorEvaluations, right.MaxSelectorEvaluations),
            MaxResponsiveImageCandidates = Minimum(left.MaxResponsiveImageCandidates, right.MaxResponsiveImageCandidates),
            MaxResponsiveImageSizesCharacters = Minimum(left.MaxResponsiveImageSizesCharacters, right.MaxResponsiveImageSizesCharacters),
            MaxSemanticMetadataCharacters = Minimum(left.MaxSemanticMetadataCharacters, right.MaxSemanticMetadataCharacters)
        };
    }

    internal void Validate() {
        ValidatePositive(MaxInputCharacters, nameof(MaxInputCharacters));
        ValidatePositive(MaxHtmlNodes, nameof(MaxHtmlNodes));
        ValidatePositive(MaxHtmlDepth, nameof(MaxHtmlDepth));
        ValidatePositive(MaxCssBytes, nameof(MaxCssBytes));
        ValidatePositive(MaxTotalCssBytes, nameof(MaxTotalCssBytes));
        ValidatePositive(MaxCssRules, nameof(MaxCssRules));
        ValidatePositive(MaxCssDeclarations, nameof(MaxCssDeclarations));
        ValidatePositive(MaxCssTokens, nameof(MaxCssTokens));
        ValidatePositive(MaxCssNestingDepth, nameof(MaxCssNestingDepth));
        ValidatePositive(MaxCssSyntaxNodes, nameof(MaxCssSyntaxNodes));
        ValidatePositive(MaxCssSelectorCharacters, nameof(MaxCssSelectorCharacters));
        ValidatePositive(MaxCssSelectorsPerRule, nameof(MaxCssSelectorsPerRule));
        ValidatePositive(MaxSelectorEvaluations, nameof(MaxSelectorEvaluations));
        ValidatePositive(MaxResponsiveImageCandidates, nameof(MaxResponsiveImageCandidates));
        ValidatePositive(MaxResponsiveImageSizesCharacters, nameof(MaxResponsiveImageSizesCharacters));
        ValidatePositive(MaxSemanticMetadataCharacters, nameof(MaxSemanticMetadataCharacters));
    }

    private static void ValidatePositive(long? value, string name) {
        if (value.HasValue && value.Value <= 0L) {
            throw new ArgumentOutOfRangeException(name, "HTML conversion limits must be positive when configured.");
        }
    }

    internal static int? Minimum(int? first, int? second) {
        if (!first.HasValue) return second;
        if (!second.HasValue) return first;
        return Math.Min(first.Value, second.Value);
    }

    private static long? Minimum(long? first, long? second) {
        if (!first.HasValue) return second;
        if (!second.HasValue) return first;
        return Math.Min(first.Value, second.Value);
    }
}

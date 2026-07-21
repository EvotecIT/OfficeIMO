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
        MaxSelectorEvaluations = 10_000_000L,
        MaxResponsiveImageCandidates = DefaultMaxResponsiveImageCandidates,
        MaxSemanticMetadataCharacters = 1024 * 1024
    };

    /// <summary>Creates compatibility-oriented limits for caller-trusted HTML.</summary>
    public static HtmlConversionLimits CreateTrustedProfile() => new HtmlConversionLimits();

    /// <summary>Maximum UTF-16 source characters, or <c>null</c> for no source-length limit.</summary>
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

    /// <summary>Maximum element/selector match attempts, or <c>null</c> for no evaluation limit.</summary>
    public long? MaxSelectorEvaluations { get; set; }

    /// <summary>Maximum responsive image candidates per source set, or <c>null</c> for no candidate limit.</summary>
    public int? MaxResponsiveImageCandidates { get; set; }

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
        MaxSelectorEvaluations = MaxSelectorEvaluations,
        MaxResponsiveImageCandidates = MaxResponsiveImageCandidates,
        MaxSemanticMetadataCharacters = MaxSemanticMetadataCharacters
    };

    internal void Validate() {
        ValidatePositive(MaxInputCharacters, nameof(MaxInputCharacters));
        ValidatePositive(MaxHtmlNodes, nameof(MaxHtmlNodes));
        ValidatePositive(MaxHtmlDepth, nameof(MaxHtmlDepth));
        ValidatePositive(MaxCssBytes, nameof(MaxCssBytes));
        ValidatePositive(MaxTotalCssBytes, nameof(MaxTotalCssBytes));
        ValidatePositive(MaxCssRules, nameof(MaxCssRules));
        ValidatePositive(MaxCssDeclarations, nameof(MaxCssDeclarations));
        ValidatePositive(MaxSelectorEvaluations, nameof(MaxSelectorEvaluations));
        ValidatePositive(MaxResponsiveImageCandidates, nameof(MaxResponsiveImageCandidates));
        ValidatePositive(MaxSemanticMetadataCharacters, nameof(MaxSemanticMetadataCharacters));
    }

    private static void ValidatePositive(long? value, string name) {
        if (value.HasValue && value.Value <= 0L) {
            throw new ArgumentOutOfRangeException(name, "HTML conversion limits must be positive when configured.");
        }
    }
}

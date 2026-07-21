using System.Collections.ObjectModel;

namespace OfficeIMO.Html;

/// <summary>
/// Stable diagnostic codes shared by OfficeIMO HTML target adapters.
/// </summary>
public static class HtmlConversionDiagnosticCodes {
    /// <summary>The expected format-specific semantic envelope was not present.</summary>
    public const string SemanticContentMissing = "SemanticContentMissing";

    /// <summary>An expected semantic table or content block was not present.</summary>
    public const string SemanticBlockMissing = "SemanticBlockMissing";

    /// <summary>A semantic scalar value could not be parsed as its declared type.</summary>
    public const string SemanticValueInvalid = "SemanticValueInvalid";

    /// <summary>An embedded resource could not be decoded.</summary>
    public const string ResourceDecodeFailed = "ResourceDecodeFailed";

    /// <summary>A resource media type is not supported by the target adapter.</summary>
    public const string ResourceTypeUnsupported = "ResourceTypeUnsupported";

    /// <summary>Content was restored through an approximate fallback.</summary>
    public const string ContentApproximated = "ContentApproximated";

    /// <summary>Content could not be represented in the target and was omitted.</summary>
    public const string ContentOmitted = "ContentOmitted";

    /// <summary>Target artifact construction failed.</summary>
    public const string ArtifactCreationFailed = "ArtifactCreationFailed";

    /// <summary>Input exceeded a target format grid or structural limit.</summary>
    public const string TargetLimitExceeded = "TargetLimitExceeded";

    /// <summary>An invalid or overlapping table span was normalized.</summary>
    public const string TableSpanInvalid = "TableSpanInvalid";

    /// <summary>HTML nesting exceeded the shared pre-analysis depth budget.</summary>
    public const string HtmlDepthLimitExceeded = "HtmlDepthLimitExceeded";

    /// <summary>One embedded stylesheet exceeded the configured byte budget.</summary>
    public const string CssSizeLimitExceeded = "CssSizeLimitExceeded";

    /// <summary>Embedded stylesheets exceeded the operation-wide byte budget.</summary>
    public const string CssTotalSizeLimitExceeded = "CssTotalSizeLimitExceeded";

    /// <summary>Active CSS rules exceeded the configured complexity budget.</summary>
    public const string CssRuleLimitExceeded = "CssRuleLimitExceeded";

    /// <summary>CSS declarations exceeded the configured complexity budget.</summary>
    public const string CssDeclarationLimitExceeded = "CssDeclarationLimitExceeded";

    /// <summary>Selector matching exceeded the configured operation-wide evaluation budget.</summary>
    public const string CssSelectorEvaluationLimitExceeded = "CssSelectorEvaluationLimitExceeded";

    /// <summary>A semantic envelope or metadata field exceeded its configured limit.</summary>
    public const string SemanticMetadataLimitExceeded = "SemanticMetadataLimitExceeded";

    /// <summary>A versioned semantic envelope declared an unsupported source or schema version.</summary>
    public const string SemanticSchemaUnsupported = "SemanticSchemaUnsupported";

    /// <summary>A legacy semantic envelope did not declare a schema version.</summary>
    public const string SemanticSchemaLegacy = "SemanticSchemaLegacy";

    /// <summary>Media filtering could not safely parse or transform an active stylesheet.</summary>
    public const string MediaFilterFailed = "MediaFilterFailed";

    /// <summary>All stable cross-adapter conversion diagnostic codes.</summary>
    public static IReadOnlyList<string> All { get; } = new ReadOnlyCollection<string>(new[] {
        SemanticContentMissing,
        SemanticBlockMissing,
        SemanticValueInvalid,
        ResourceDecodeFailed,
        ResourceTypeUnsupported,
        ContentApproximated,
        ContentOmitted,
        ArtifactCreationFailed,
        TargetLimitExceeded,
        TableSpanInvalid,
        HtmlDepthLimitExceeded,
        CssSizeLimitExceeded,
        CssTotalSizeLimitExceeded,
        CssRuleLimitExceeded,
        CssDeclarationLimitExceeded,
        CssSelectorEvaluationLimitExceeded,
        SemanticMetadataLimitExceeded,
        SemanticSchemaUnsupported,
        SemanticSchemaLegacy,
        MediaFilterFailed
    });
}

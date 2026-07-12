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
}

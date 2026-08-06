namespace OfficeIMO;

/// <summary>Open XML SDK compatibility behavior used when a package is opened.</summary>
public enum OfficeOpenXmlCompatibilityLevel {
    /// <summary>Use the SDK's current default behavior.</summary>
    Default,

    /// <summary>Use behavior compatible with Open XML SDK 2.20.</summary>
    Version220,

    /// <summary>Use behavior compatible with Open XML SDK 3.x.</summary>
    Version30
}

/// <summary>Controls how markup-compatibility elements are processed while loading an Open XML package.</summary>
public enum OfficeOpenXmlMarkupCompatibilityMode {
    /// <summary>Do not process markup-compatibility elements.</summary>
    NoProcess,

    /// <summary>Process markup compatibility only for parts loaded by the caller.</summary>
    ProcessLoadedPartsOnly,

    /// <summary>Process markup compatibility for every package part.</summary>
    ProcessAllParts
}

/// <summary>Microsoft Office generation used for Open XML validation and markup compatibility.</summary>
public enum OfficeOpenXmlFileFormatVersion {
    /// <summary>Office 2007.</summary>
    Office2007,

    /// <summary>Office 2010.</summary>
    Office2010,

    /// <summary>Office 2013.</summary>
    Office2013,

    /// <summary>Office 2016.</summary>
    Office2016,

    /// <summary>Office 2019.</summary>
    Office2019,

    /// <summary>Office 2021.</summary>
    Office2021,

    /// <summary>Microsoft 365.</summary>
    Microsoft365
}

/// <summary>Low-level but SDK-independent controls for opening an Open XML package.</summary>
public sealed class OfficeOpenXmlLoadSettings {
    /// <summary>Gets or sets the Open XML SDK compatibility behavior.</summary>
    public OfficeOpenXmlCompatibilityLevel CompatibilityLevel { get; set; } = OfficeOpenXmlCompatibilityLevel.Version30;

    /// <summary>Gets or sets how markup-compatibility elements are processed.</summary>
    public OfficeOpenXmlMarkupCompatibilityMode MarkupCompatibilityMode { get; set; } = OfficeOpenXmlMarkupCompatibilityMode.NoProcess;

    /// <summary>Gets or sets the Office generation targeted by markup-compatibility processing.</summary>
    public OfficeOpenXmlFileFormatVersion MarkupCompatibilityTargetVersion { get; set; } = OfficeOpenXmlFileFormatVersion.Office2007;

    /// <summary>
    /// Gets or sets the maximum number of characters allowed in one XML part.
    /// A value of zero uses the Open XML SDK default.
    /// </summary>
    public long MaxCharactersInPart { get; set; }
}

/// <summary>Classifies an Open XML package validation finding.</summary>
public enum OfficeOpenXmlValidationErrorType {
    /// <summary>Schema validation failure.</summary>
    Schema,

    /// <summary>Semantic validation failure.</summary>
    Semantic,

    /// <summary>Package structure validation failure.</summary>
    Package,

    /// <summary>Markup-compatibility validation failure.</summary>
    MarkupCompatibility
}

/// <summary>SDK-independent description of one Open XML package validation finding.</summary>
public sealed class OfficeOpenXmlValidationError {
    internal OfficeOpenXmlValidationError(
        string? id,
        OfficeOpenXmlValidationErrorType errorType,
        string? description,
        string? path,
        string? partUri,
        string? nodeName,
        string? relatedPartUri,
        string? relatedNodeName) {
        Id = id ?? string.Empty;
        ErrorType = errorType;
        Description = description ?? string.Empty;
        Path = path;
        PartUri = partUri;
        NodeName = nodeName;
        RelatedPartUri = relatedPartUri;
        RelatedNodeName = relatedNodeName;
    }

    /// <summary>Gets the stable validator identifier, when supplied.</summary>
    public string Id { get; }

    /// <summary>Gets the validation finding category.</summary>
    public OfficeOpenXmlValidationErrorType ErrorType { get; }

    /// <summary>Gets the human-readable validation message.</summary>
    public string Description { get; }

    /// <summary>Gets the XPath-like package location, when supplied.</summary>
    public string? Path { get; }

    /// <summary>Gets the package part URI, when supplied.</summary>
    public string? PartUri { get; }

    /// <summary>Gets the affected element name, when supplied.</summary>
    public string? NodeName { get; }

    /// <summary>Gets the related package part URI, when supplied.</summary>
    public string? RelatedPartUri { get; }

    /// <summary>Gets the related element name, when supplied.</summary>
    public string? RelatedNodeName { get; }
}

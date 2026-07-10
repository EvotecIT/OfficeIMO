namespace OfficeIMO.OpenDocument;

/// <summary>Describes how a source feature was handled during an explicit format conversion.</summary>
public enum OdfConversionMappingStatus {
    /// <summary>The supported semantic value was transferred directly.</summary>
    Converted,
    /// <summary>The feature was transferred through a documented approximation.</summary>
    Approximated,
    /// <summary>The feature was intentionally omitted from the target.</summary>
    Skipped,
    /// <summary>The target adapter does not currently support the source feature.</summary>
    Unsupported
}

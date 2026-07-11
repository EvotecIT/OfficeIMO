namespace OfficeIMO.OpenDocument;

/// <summary>Describes how the current engine handles a detected document feature.</summary>
public enum OdfFeatureSupport {
    /// <summary>The feature has a typed editing surface.</summary>
    Editable,
    /// <summary>The feature is inspectable but not editable.</summary>
    Inspected,
    /// <summary>The feature is retained without typed interpretation.</summary>
    Preserved,
    /// <summary>The feature cannot be retained across the requested operation.</summary>
    Unsupported
}

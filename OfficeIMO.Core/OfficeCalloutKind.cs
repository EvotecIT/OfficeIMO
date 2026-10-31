namespace OfficeIMO;

/// <summary>The intent of a document callout, resolved through its theme palette.</summary>
public enum OfficeCalloutKind {
    /// <summary>General information.</summary>
    Info,
    /// <summary>A successful result or helpful tip.</summary>
    Success,
    /// <summary>A condition requiring attention.</summary>
    Warning,
    /// <summary>An error or critical condition.</summary>
    Danger
}

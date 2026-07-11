namespace OfficeIMO.OpenDocument;

/// <summary>Controls the version and extension policy used when saving.</summary>
public enum OdfCompatibilityProfile {
    /// <summary>Write conforming ODF 1.4 markup for newly supported features.</summary>
    Odf14,
    /// <summary>Write ODF 1.3-compatible markup where supported.</summary>
    Odf13,
    /// <summary>Retain the source version and extension markup where possible.</summary>
    PreserveSource
}

namespace OfficeIMO.OpenDocument;

/// <summary>Controls saving when a mutation would invalidate document signatures.</summary>
public enum OdfSignatureHandling {
    /// <summary>Reject saving a changed signed document.</summary>
    RejectInvalidation,
    /// <summary>Remove signature entries invalidated by the requested changes.</summary>
    RemoveInvalidated
}

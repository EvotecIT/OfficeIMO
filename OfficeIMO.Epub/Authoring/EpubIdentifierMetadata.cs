namespace OfficeIMO.Epub;

/// <summary>Identifier forms supported by typed bibliographic authoring.</summary>
public enum EpubIdentifierKind {
    /// <summary>An opaque identifier; no identifier-type refinement is emitted.</summary>
    Unspecified,
    /// <summary>A historical ten-character ISBN, with its modulo-11 check digit.</summary>
    Isbn10,
    /// <summary>A thirteen-digit ISBN, with prefix and modulo-10 check digit.</summary>
    Isbn13,
    /// <summary>A DOI name beginning with 10., without a resolver URL prefix.</summary>
    Doi
}

/// <summary>An additional publication identifier. This does not select or replace the package identity.</summary>
public sealed class EpubIdentifierMetadata {
    /// <summary>
    /// Identifier value. ISBNs accept ASCII spaces, hyphens and an optional urn:isbn: prefix;
    /// output uses a compact ISBN URN. DOI names and opaque identifiers retain their supplied text.
    /// </summary>
    public string Value { get; set; } = string.Empty;
    /// <summary>Identifier form; assigned ranges and registration are not verified.</summary>
    public EpubIdentifierKind Kind { get; set; }
}

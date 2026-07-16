namespace OfficeIMO.Epub;

/// <summary>Describes why an EPUB URL reference could not be resolved.</summary>
public enum EpubReferenceError {
    /// <summary>The reference resolved successfully.</summary>
    None,

    /// <summary>The reference value was empty.</summary>
    Empty,

    /// <summary>The container base path was not a safe archive path.</summary>
    InvalidBasePath,

    /// <summary>The reference contains a control character.</summary>
    ControlCharacter,

    /// <summary>The reference uses the prohibited <c>file:</c> URL scheme.</summary>
    FileUrl,

    /// <summary>The reference contains an invalid or ambiguous path encoding.</summary>
    InvalidPath,

    /// <summary>The reference attempts to traverse above the EPUB container root.</summary>
    EscapesContainer
}

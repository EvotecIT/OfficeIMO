namespace OfficeIMO.Rtf;

/// <summary>
/// RTF document character set declaration.
/// </summary>
public enum RtfDocumentCharacterSet {
    /// <summary>ANSI character set represented by <c>\ansi</c>.</summary>
    Ansi,

    /// <summary>Macintosh character set represented by <c>\mac</c>.</summary>
    Mac,

    /// <summary>IBM PC code page 437 character set represented by <c>\pc</c>.</summary>
    Pc,

    /// <summary>IBM PC code page 850 character set represented by <c>\pca</c>.</summary>
    Pca
}

namespace OfficeIMO.Epub;

/// <summary>Classifies encryption declarations used by EPUB containers.</summary>
public enum EpubEncryptionKind {
    /// <summary>The declaration does not provide a recognized algorithm.</summary>
    Unknown,

    /// <summary>IDPF EPUB font obfuscation.</summary>
    IdpfFontObfuscation,

    /// <summary>Legacy Adobe font obfuscation.</summary>
    AdobeFontObfuscation,

    /// <summary>Encryption that requires an external decryption or DRM capability.</summary>
    Encryption
}

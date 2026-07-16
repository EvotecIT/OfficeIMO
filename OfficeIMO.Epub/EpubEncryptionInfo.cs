namespace OfficeIMO.Epub;

/// <summary>Describes an encrypted or obfuscated resource declared by META-INF/encryption.xml.</summary>
public sealed class EpubEncryptionInfo {
    /// <summary>Normalized archive path of the affected resource.</summary>
    public string Path { get; internal set; } = string.Empty;

    /// <summary>Declared XML Encryption algorithm URI.</summary>
    public string? Algorithm { get; internal set; }

    /// <summary>Classified encryption or obfuscation kind.</summary>
    public EpubEncryptionKind Kind { get; internal set; }

    /// <summary>Whether the declaration is a recognized EPUB font-obfuscation algorithm.</summary>
    public bool IsFontObfuscation =>
        Kind == EpubEncryptionKind.IdpfFontObfuscation ||
        Kind == EpubEncryptionKind.AdobeFontObfuscation;

    /// <summary>Whether the resource requires a decryption capability not provided by OfficeIMO.Epub.</summary>
    public bool RequiresDecryption => Kind == EpubEncryptionKind.Encryption || Kind == EpubEncryptionKind.Unknown;
}

namespace OfficeIMO.Security;

/// <summary>Content-encryption algorithms offered by the neutral CMS envelope service.</summary>
public enum CmsContentEncryptionAlgorithm {
    /// <summary>AES-128 in CBC mode.</summary>
    Aes128Cbc = 0,
    /// <summary>AES-192 in CBC mode.</summary>
    Aes192Cbc = 1,
    /// <summary>AES-256 in CBC mode. This is the default.</summary>
    Aes256Cbc = 2
}

/// <summary>Safety and algorithm options for CMS EnvelopedData.</summary>
public sealed class CmsEnvelopeOptions {
    /// <summary>Content-encryption algorithm. Defaults to AES-256-CBC.</summary>
    public CmsContentEncryptionAlgorithm ContentEncryptionAlgorithm { get; set; } =
        CmsContentEncryptionAlgorithm.Aes256Cbc;
    /// <summary>Maximum plaintext bytes accepted or returned. Defaults to 512 MiB.</summary>
    public long MaxContentBytes { get; set; } = 512L * 1024 * 1024;
    /// <summary>Maximum encoded envelope bytes accepted. Defaults to 512 MiB.</summary>
    public long MaxEncodedBytes { get; set; } = 512L * 1024 * 1024;
    /// <summary>Maximum recipient count. Defaults to 64.</summary>
    public int MaxRecipients { get; set; } = 64;
}

/// <summary>Result of attempting to decrypt CMS EnvelopedData for one certificate.</summary>
public sealed class CmsDecryptionResult {
    internal CmsDecryptionResult(
        bool parsed,
        bool decrypted,
        byte[]? content,
        string? contentEncryptionAlgorithmOid,
        string? keyEncryptionAlgorithmOid,
        IReadOnlyList<SecurityFinding> findings) {
        Parsed = parsed;
        Decrypted = decrypted;
        Content = content;
        ContentEncryptionAlgorithmOid = contentEncryptionAlgorithmOid;
        KeyEncryptionAlgorithmOid = keyEncryptionAlgorithmOid;
        Findings = findings;
    }

    /// <summary>Whether the EnvelopedData container was decoded.</summary>
    public bool Parsed { get; }
    /// <summary>Whether a matching recipient was successfully decrypted.</summary>
    public bool Decrypted { get; }
    /// <summary>Cloned plaintext content when decryption succeeded.</summary>
    public byte[]? Content { get; }
    /// <summary>CMS content-encryption algorithm object identifier.</summary>
    public string? ContentEncryptionAlgorithmOid { get; }
    /// <summary>Recipient key-encryption algorithm object identifier.</summary>
    public string? KeyEncryptionAlgorithmOid { get; }
    /// <summary>Structured processing findings.</summary>
    public IReadOnlyList<SecurityFinding> Findings { get; }
}

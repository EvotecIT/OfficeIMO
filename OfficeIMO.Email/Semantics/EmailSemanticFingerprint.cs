namespace OfficeIMO.Email;

/// <summary>Versioned digest of a canonical email semantic projection.</summary>
public sealed class EmailSemanticFingerprint {
    private readonly byte[] _digest;

    internal EmailSemanticFingerprint(int schemaVersion, string algorithm, byte[] digest,
        EmailSemanticComparisonProfile profile, int recipientCount, int attachmentCount,
        long attachmentBytesHashed, int entryCount) {
        SchemaVersion = schemaVersion;
        Algorithm = algorithm;
        _digest = (byte[])digest.Clone();
        Profile = profile;
        RecipientCount = recipientCount;
        AttachmentCount = attachmentCount;
        AttachmentBytesHashed = attachmentBytesHashed;
        EntryCount = entryCount;
    }

    /// <summary>Canonical schema version. Fingerprints with different versions are not comparable.</summary>
    public int SchemaVersion { get; }

    /// <summary>Digest algorithm, currently SHA-256 or HMAC-SHA-256.</summary>
    public string Algorithm { get; }

    /// <summary>Semantic comparison profile used to produce the digest.</summary>
    public EmailSemanticComparisonProfile Profile { get; }

    /// <summary>Digest bytes. A defensive copy is returned.</summary>
    public byte[] Digest => (byte[])_digest.Clone();

    /// <summary>Uppercase hexadecimal digest.</summary>
    public string HexDigest => BitConverter.ToString(_digest).Replace("-", string.Empty);

    /// <summary>Number of projected recipients.</summary>
    public int RecipientCount { get; }

    /// <summary>Number of projected attachments, including nested attachments.</summary>
    public int AttachmentCount { get; }

    /// <summary>Decoded attachment bytes included in the digest.</summary>
    public long AttachmentBytesHashed { get; }

    /// <summary>Number of canonical leaf entries included in the digest.</summary>
    public int EntryCount { get; }
}

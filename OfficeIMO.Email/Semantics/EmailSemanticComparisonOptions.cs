namespace OfficeIMO.Email;

/// <summary>Immutable limits and policies for semantic fingerprinting and comparison.</summary>
public sealed class EmailSemanticComparisonOptions {
    private readonly byte[]? _digestKey;

    /// <summary>Default migration-verification policy.</summary>
    public static EmailSemanticComparisonOptions Default { get; } = new EmailSemanticComparisonOptions();

    /// <summary>Creates a semantic comparison policy.</summary>
    /// <param name="profile">Representation profile to compare.</param>
    /// <param name="digestKey">
    /// Optional HMAC key. Supply a caller-owned random secret before persisting fingerprints of private content.
    /// The key is copied and is never exposed by the resulting fingerprint.
    /// </param>
    /// <param name="includeAttachmentContent">Whether decoded attachment content participates in comparison.</param>
    /// <param name="maxAttachmentBytes">Maximum decoded bytes hashed for one attachment.</param>
    /// <param name="maxTotalAttachmentBytes">Maximum decoded attachment bytes hashed across one root document.</param>
    /// <param name="maxEmbeddedMessageDepth">Maximum embedded-message recursion depth.</param>
    /// <param name="maxDifferences">Maximum detailed differences retained by a comparison report.</param>
    public EmailSemanticComparisonOptions(
        EmailSemanticComparisonProfile profile = EmailSemanticComparisonProfile.Migration,
        byte[]? digestKey = null,
        bool includeAttachmentContent = true,
        long maxAttachmentBytes = 8L * 1024 * 1024 * 1024,
        long maxTotalAttachmentBytes = 64L * 1024 * 1024 * 1024,
        int maxEmbeddedMessageDepth = 32,
        int maxDifferences = 10_000) {
        if (maxAttachmentBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxAttachmentBytes));
        if (maxTotalAttachmentBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxTotalAttachmentBytes));
        if (maxEmbeddedMessageDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxEmbeddedMessageDepth));
        if (maxDifferences <= 0) throw new ArgumentOutOfRangeException(nameof(maxDifferences));
        if (digestKey != null && digestKey.Length < 16) {
            throw new ArgumentException("A semantic HMAC key must contain at least 16 bytes.", nameof(digestKey));
        }

        Profile = profile;
        _digestKey = digestKey == null ? null : (byte[])digestKey.Clone();
        IncludeAttachmentContent = includeAttachmentContent;
        MaxAttachmentBytes = maxAttachmentBytes;
        MaxTotalAttachmentBytes = maxTotalAttachmentBytes;
        MaxEmbeddedMessageDepth = maxEmbeddedMessageDepth;
        MaxDifferences = maxDifferences;
    }

    /// <summary>Representation profile being compared.</summary>
    public EmailSemanticComparisonProfile Profile { get; }

    /// <summary>True when fingerprints use HMAC-SHA-256 rather than unkeyed SHA-256.</summary>
    public bool UsesKeyedDigest => _digestKey != null;

    /// <summary>Whether decoded attachment content participates in comparison.</summary>
    public bool IncludeAttachmentContent { get; }

    /// <summary>Maximum decoded bytes hashed for one attachment.</summary>
    public long MaxAttachmentBytes { get; }

    /// <summary>Maximum decoded attachment bytes hashed across one root document.</summary>
    public long MaxTotalAttachmentBytes { get; }

    /// <summary>Maximum embedded-message recursion depth.</summary>
    public int MaxEmbeddedMessageDepth { get; }

    /// <summary>Maximum detailed differences retained by one comparison report.</summary>
    public int MaxDifferences { get; }

    internal byte[]? CopyDigestKey() => _digestKey == null ? null : (byte[])_digestKey.Clone();
}

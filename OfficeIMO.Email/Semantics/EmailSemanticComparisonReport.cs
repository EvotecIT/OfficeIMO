namespace OfficeIMO.Email;

/// <summary>Privacy-safe result of comparing two canonical email projections.</summary>
public sealed class EmailSemanticComparisonReport {
    internal EmailSemanticComparisonReport(EmailSemanticFingerprint source,
        EmailSemanticFingerprint destination, IReadOnlyList<EmailSemanticDifference> differences,
        bool differencesTruncated) {
        Source = source;
        Destination = destination;
        Differences = differences;
        DifferencesTruncated = differencesTruncated;
    }

    /// <summary>Source fingerprint.</summary>
    public EmailSemanticFingerprint Source { get; }

    /// <summary>Destination fingerprint.</summary>
    public EmailSemanticFingerprint Destination { get; }

    /// <summary>Detailed differences, capped by the configured limit.</summary>
    public IReadOnlyList<EmailSemanticDifference> Differences { get; }

    /// <summary>True when additional differences existed beyond the configured detail limit.</summary>
    public bool DifferencesTruncated { get; }

    /// <summary>True when both semantic projections match exactly under the selected profile.</summary>
    public bool IsMatch => Differences.Count == 0 && !DifferencesTruncated &&
        Source.HexDigest == Destination.HexDigest;
}

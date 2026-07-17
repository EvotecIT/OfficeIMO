namespace OfficeIMO.Email;

/// <summary>Creates versioned semantic fingerprints and privacy-safe comparison reports.</summary>
public static class EmailSemanticComparer {
    /// <summary>Current canonical semantic projection schema.</summary>
    public const int CurrentSchemaVersion = 1;

    /// <summary>Creates a semantic fingerprint, streaming deferred attachment content.</summary>
    public static EmailSemanticFingerprint CreateFingerprint(EmailDocument document,
        EmailSemanticComparisonOptions? options = null,
        CancellationToken cancellationToken = default) {
        EmailSemanticSnapshot snapshot = new EmailSemanticSnapshotBuilder(
            options ?? EmailSemanticComparisonOptions.Default).Build(document, cancellationToken);
        return snapshot.Fingerprint;
    }

    /// <summary>Asynchronously creates a semantic fingerprint, including deferred attachment streams.</summary>
    public static async Task<EmailSemanticFingerprint> CreateFingerprintAsync(EmailDocument document,
        EmailSemanticComparisonOptions? options = null,
        CancellationToken cancellationToken = default) {
        EmailSemanticSnapshot snapshot = await new EmailSemanticSnapshotBuilder(
            options ?? EmailSemanticComparisonOptions.Default).BuildAsync(document, cancellationToken)
            .ConfigureAwait(false);
        return snapshot.Fingerprint;
    }

    /// <summary>Compares two documents without returning their values or component digests.</summary>
    public static EmailSemanticComparisonReport Compare(EmailDocument source,
        EmailDocument destination, EmailSemanticComparisonOptions? options = null,
        CancellationToken cancellationToken = default) {
        var effective = options ?? EmailSemanticComparisonOptions.Default;
        EmailSemanticSnapshot sourceSnapshot = new EmailSemanticSnapshotBuilder(effective)
            .Build(source, cancellationToken);
        EmailSemanticSnapshot destinationSnapshot = new EmailSemanticSnapshotBuilder(effective)
            .Build(destination, cancellationToken);
        return CompareSnapshots(sourceSnapshot, destinationSnapshot, effective.MaxDifferences);
    }

    /// <summary>Asynchronously compares two documents, including deferred attachment streams.</summary>
    public static async Task<EmailSemanticComparisonReport> CompareAsync(EmailDocument source,
        EmailDocument destination, EmailSemanticComparisonOptions? options = null,
        CancellationToken cancellationToken = default) {
        var effective = options ?? EmailSemanticComparisonOptions.Default;
        EmailSemanticSnapshot sourceSnapshot = await new EmailSemanticSnapshotBuilder(effective)
            .BuildAsync(source, cancellationToken).ConfigureAwait(false);
        EmailSemanticSnapshot destinationSnapshot = await new EmailSemanticSnapshotBuilder(effective)
            .BuildAsync(destination, cancellationToken).ConfigureAwait(false);
        return CompareSnapshots(sourceSnapshot, destinationSnapshot, effective.MaxDifferences);
    }

    private static EmailSemanticComparisonReport CompareSnapshots(EmailSemanticSnapshot source,
        EmailSemanticSnapshot destination, int maximumDifferences) {
        var differences = new List<EmailSemanticDifference>();
        bool truncated = false;
        string[] paths = source.Entries.Keys.Concat(destination.Entries.Keys)
            .Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToArray();
        foreach (string path in paths) {
            source.Entries.TryGetValue(path, out EmailSemanticEntry? sourceEntry);
            destination.Entries.TryGetValue(path, out EmailSemanticEntry? destinationEntry);
            EmailSemanticDifferenceKind? kind = sourceEntry == null
                ? EmailSemanticDifferenceKind.UnexpectedInDestination
                : destinationEntry == null
                    ? EmailSemanticDifferenceKind.MissingFromDestination
                    : FixedTimeEquals(sourceEntry.Digest, destinationEntry.Digest) &&
                      sourceEntry.Length == destinationEntry.Length
                        ? (EmailSemanticDifferenceKind?)null
                        : EmailSemanticDifferenceKind.Changed;
            if (!kind.HasValue) continue;
            if (differences.Count >= maximumDifferences) {
                truncated = true;
                continue;
            }
            differences.Add(new EmailSemanticDifference(path, kind.Value,
                sourceEntry?.Length, destinationEntry?.Length));
        }
        return new EmailSemanticComparisonReport(source.Fingerprint, destination.Fingerprint,
            differences.AsReadOnly(), truncated);
    }

    private static bool FixedTimeEquals(byte[] left, byte[] right) {
        if (left.Length != right.Length) return false;
        int difference = 0;
        for (int index = 0; index < left.Length; index++) difference |= left[index] ^ right[index];
        return difference == 0;
    }
}

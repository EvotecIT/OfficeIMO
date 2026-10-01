namespace OfficeIMO.Email.Store;

/// <summary>Limits for a read-only analysis of regular items in an open archive.</summary>
public sealed class EmailArchiveAnalysisOptions {
    /// <summary>Creates finite scan, projection and result bounds. No attachment payload is requested.</summary>
    public EmailArchiveAnalysisOptions(int maxItems = 10000, long maxDecodedPropertyBytesPerItem = 8L * 1024 * 1024,
        int maxDuplicateGroups = 100, int maxItemsPerDuplicateGroup = 100,
        int maxLargeAttachments = 50, long largeAttachmentThresholdBytes = 1024 * 1024,
        int maxDistributionBuckets = 1000, int maxDiagnostics = 100) {
        if (maxItems <= 0 || maxItems > 100000) throw new ArgumentOutOfRangeException(nameof(maxItems));
        if (maxDecodedPropertyBytesPerItem <= 0 || maxDecodedPropertyBytesPerItem > 64L * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(maxDecodedPropertyBytesPerItem));
        if (maxDuplicateGroups <= 0 || maxDuplicateGroups > 1000) throw new ArgumentOutOfRangeException(nameof(maxDuplicateGroups));
        if (maxItemsPerDuplicateGroup <= 0 || maxItemsPerDuplicateGroup > 1000) throw new ArgumentOutOfRangeException(nameof(maxItemsPerDuplicateGroup));
        if (maxLargeAttachments <= 0 || maxLargeAttachments > 1000) throw new ArgumentOutOfRangeException(nameof(maxLargeAttachments));
        if (largeAttachmentThresholdBytes < 0) throw new ArgumentOutOfRangeException(nameof(largeAttachmentThresholdBytes));
        if (maxDistributionBuckets <= 0 || maxDistributionBuckets > 10000) throw new ArgumentOutOfRangeException(nameof(maxDistributionBuckets));
        if (maxDiagnostics <= 0 || maxDiagnostics > 1000) throw new ArgumentOutOfRangeException(nameof(maxDiagnostics));
        MaxItems = maxItems; MaxDecodedPropertyBytesPerItem = maxDecodedPropertyBytesPerItem;
        MaxDuplicateGroups = maxDuplicateGroups; MaxItemsPerDuplicateGroup = maxItemsPerDuplicateGroup;
        MaxLargeAttachments = maxLargeAttachments; LargeAttachmentThresholdBytes = largeAttachmentThresholdBytes;
        MaxDistributionBuckets = maxDistributionBuckets; MaxDiagnostics = maxDiagnostics;
    }
    /// <summary>Maximum regular references projected across all folders.</summary>
    public int MaxItems { get; }
    /// <summary>Maximum decoded properties for one selected item; the session policy can narrow it.</summary>
    public long MaxDecodedPropertyBytesPerItem { get; }
    /// <summary>Maximum candidate groups retained in the report.</summary>
    public int MaxDuplicateGroups { get; }
    /// <summary>Maximum source IDs retained per candidate group.</summary>
    public int MaxItemsPerDuplicateGroup { get; }
    /// <summary>Maximum large-attachment rows retained.</summary>
    public int MaxLargeAttachments { get; }
    /// <summary>Inclusive declared attachment-size threshold.</summary>
    public long LargeAttachmentThresholdBytes { get; }
    /// <summary>Maximum distinct buckets retained in each folder/month distribution.</summary>
    public int MaxDistributionBuckets { get; }
    /// <summary>Maximum diagnostic samples retained; total counts remain explicit.</summary>
    public int MaxDiagnostics { get; }
}

/// <summary>One bounded distribution bucket.</summary>
public sealed class EmailArchiveDistributionBucket {
    internal EmailArchiveDistributionBucket(string key, int count) { Key = key; Count = count; }
    /// <summary>Folder ID or UTC yyyy-MM month.</summary>
    public string Key { get; }
    /// <summary>Scanned references in this bucket.</summary>
    public int Count { get; }
}

/// <summary>Items with matching projected semantics and attachment metadata. Payload equality is unverified.</summary>
public sealed class EmailArchiveDuplicateCandidateGroup {
    internal EmailArchiveDuplicateCandidateGroup(IReadOnlyList<string> ids, int count, long bytes) { ItemIds = ids; ItemCount = count; EstimatedDeclaredBytes = bytes; }
    /// <summary>Bounded source IDs, in enumeration order. No item is selected for deletion.</summary>
    public IReadOnlyList<string> ItemIds { get; }
    /// <summary>Total candidate members, including IDs omitted by the result bound.</summary>
    public int ItemCount { get; }
    /// <summary>Whether some member IDs were omitted.</summary>
    public bool ItemIdsTruncated => ItemIds.Count < ItemCount;
    /// <summary>Sum of nonnegative declared item sizes minus the largest known member size. Unknown sizes contribute zero.</summary>
    public long EstimatedDeclaredBytes { get; }
}

/// <summary>A large attachment identified from available metadata, without opening its payload.</summary>
public sealed class EmailArchiveLargeAttachment {
    internal EmailArchiveLargeAttachment(string id, int index, string? name, long bytes, bool embedded) {
        ItemId = id; AttachmentIndex = index; FileName = name; DeclaredBytes = bytes; IsEmbeddedItem = embedded;
    }
    /// <summary>Owning source item ID.</summary>
    public string ItemId { get; }
    /// <summary>Zero-based attachment index on the selected item.</summary>
    public int AttachmentIndex { get; }
    /// <summary>Filename preview, limited to 256 UTF-16 units without splitting a surrogate pair.</summary>
    public string? FileName { get; }
    /// <summary>Available declared size; this is not measured payload I/O.</summary>
    public long DeclaredBytes { get; }
    /// <summary>Whether metadata classifies the attachment as an embedded item.</summary>
    public bool IsEmbeddedItem { get; }
}

/// <summary>Source-bound statistics and candidate evidence for the selected regular-item enumeration.</summary>
public sealed class EmailArchiveAnalysisReport {
    internal EmailArchiveAnalysisReport(string fingerprint, int scanned, int projected, int failed, int partial, bool scanLimit,
        IReadOnlyList<EmailArchiveDistributionBucket> folders, int otherFolders, IReadOnlyList<EmailArchiveDistributionBucket> months,
        int otherMonths, int unknownDates, IReadOnlyList<EmailArchiveDuplicateCandidateGroup> groups, int groupCount,
        long estimatedBytes, IReadOnlyList<EmailArchiveLargeAttachment> attachments, int largeCount, int unknownSizes,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics, int diagnosticCount, bool sourceWarnings) {
        SourceFingerprint = fingerprint; ItemsScanned = scanned; ItemsProjected = projected; ItemsFailed = failed;
        PartialItemsExcludedFromCandidates = partial; StoppedAtItemLimit = scanLimit;
        Folders = folders; ItemsInOtherFolders = otherFolders; UtcMonths = months; ItemsInOtherMonths = otherMonths;
        ItemsWithoutDate = unknownDates; DuplicateCandidates = groups; DuplicateCandidateGroupCount = groupCount;
        EstimatedCandidateDeclaredBytes = estimatedBytes; LargeAttachments = attachments; LargeAttachmentCount = largeCount;
        AttachmentsWithoutPositiveSize = unknownSizes; Diagnostics = diagnostics; DiagnosticCount = diagnosticCount; HasSourceWarnings = sourceWarnings;
    }
    /// <summary>SHA-256 of the complete persisted source, rechecked after analysis. Computing this performs source I/O.</summary>
    public string SourceFingerprint { get; }
    /// <summary>References considered, including failed projections.</summary>
    public int ItemsScanned { get; }
    /// <summary>Successfully projected items.</summary>
    public int ItemsProjected { get; }
    /// <summary>Items not fully analyzed. This can include projected items whose semantic fingerprinting failed.</summary>
    public int ItemsFailed { get; }
    /// <summary>Header-only or potentially partial local items omitted from semantic candidate matching.</summary>
    public int PartialItemsExcludedFromCandidates { get; }
    /// <summary>Whether another regular reference existed beyond the scan limit.</summary>
    public bool StoppedAtItemLimit { get; }
    /// <summary>Whether the session's selected regular-reference enumeration was exhausted; this does not prove source-catalog completeness.</summary>
    public bool ExhaustedSelectedReferences => !StoppedAtItemLimit;
    /// <summary>Whether the source reported compatibility, recovery or completeness warnings.</summary>
    public bool HasSourceWarnings { get; }
    /// <summary>Folder distribution for scanned references.</summary>
    public IReadOnlyList<EmailArchiveDistributionBucket> Folders { get; }
    /// <summary>Scanned references belonging to folder buckets omitted by the bound.</summary>
    public int ItemsInOtherFolders { get; }
    /// <summary>UTC month distribution, preferring sent time then received time.</summary>
    public IReadOnlyList<EmailArchiveDistributionBucket> UtcMonths { get; }
    /// <summary>Dated items belonging to month buckets omitted by the bound.</summary>
    public int ItemsInOtherMonths { get; }
    /// <summary>Scanned items whose date was unavailable, including failed projections.</summary>
    public int ItemsWithoutDate { get; }
    /// <summary>Bounded candidate groups. Attachment content, embedded content and unloaded extended properties are not verified.</summary>
    public IReadOnlyList<EmailArchiveDuplicateCandidateGroup> DuplicateCandidates { get; }
    /// <summary>Total candidate groups detected in the scanned projection.</summary>
    public int DuplicateCandidateGroupCount { get; }
    /// <summary>Whether some candidate groups are absent from the sample.</summary>
    public bool DuplicateGroupsTruncated => DuplicateCandidates.Count < DuplicateCandidateGroupCount;
    /// <summary>Declared-byte estimate across candidate groups. It is neither validated duplicate savings nor a PST compaction estimate.</summary>
    public long EstimatedCandidateDeclaredBytes { get; }
    /// <summary>Largest available attachment metadata rows in descending size order.</summary>
    public IReadOnlyList<EmailArchiveLargeAttachment> LargeAttachments { get; }
    /// <summary>Total metadata rows meeting the size threshold.</summary>
    public int LargeAttachmentCount { get; }
    /// <summary>Rows with a zero, negative or unknown available length. This can include empty attachments.</summary>
    public int AttachmentsWithoutPositiveSize { get; }
    /// <summary>Bounded diagnostic sample, without exception content or body previews.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    /// <summary>Total source and analysis diagnostics.</summary>
    public int DiagnosticCount { get; }
}

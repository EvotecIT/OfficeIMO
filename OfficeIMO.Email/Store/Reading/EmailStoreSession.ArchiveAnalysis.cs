namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Analyzes regular archive items without changing the source or requesting attachment payloads.
    /// Semantic matches are candidates only; use approved maintenance operations for any later rewrite.
    /// </summary>
    public EmailArchiveAnalysisReport AnalyzeArchive(EmailArchiveAnalysisOptions? options = null,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed(); var policy = options ?? new EmailArchiveAnalysisOptions();
        string fingerprint = GetDurableSourceFingerprint(cancellationToken);
        var folders = new Dictionary<string, int>(StringComparer.Ordinal);
        var months = new Dictionary<string, int>(StringComparer.Ordinal);
        var candidates = new Dictionary<string, ArchiveCandidateAccumulator>(StringComparer.Ordinal);
        var large = new List<EmailArchiveLargeAttachment>(); var diagnostics = new List<EmailStoreDiagnostic>();
        var semanticPolicy = new EmailSemanticComparisonOptions(EmailSemanticComparisonProfile.Deduplication,
            includeAttachmentContent: false, maxEmbeddedMessageDepth: 0, includeEmbeddedMessageContent: false);
        var readPolicy = new EmailStoreItemReadOptions(EmailStoreItemReadParts.Metadata | EmailStoreItemReadParts.Bodies |
            EmailStoreItemReadParts.Recipients | EmailStoreItemReadParts.AttachmentMetadata,
            policy.MaxDecodedPropertyBytesPerItem, preferStreamingAttachmentContent: true);
        int scanned = 0, projected = 0, failed = 0, partial = 0, otherFolders = 0, otherMonths = 0,
            unknownDates = 0, largeCount = 0, unknownSizes = 0, diagnosticCount = 0;
        bool scanLimit = false;
        foreach (EmailStoreItemReference reference in EnumerateItems(new EmailStoreEnumerationOptions(maxItems: policy.MaxItems + 1), cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (scanned == policy.MaxItems) { scanLimit = true; break; }
            scanned++;
            if (!AddBucket(folders, reference.FolderId)) otherFolders++;
            bool dateCounted = false;
            try {
                EmailStoreItem item = ReadItem(reference, readPolicy, cancellationToken);
                EmailDocument document = item.Document;
                projected++;
                DateTimeOffset? date = document.Date ?? document.ReceivedDate;
                if (!date.HasValue) unknownDates++;
                else if (!AddBucket(months, date.Value.UtcDateTime.ToString("yyyy-MM", CultureInfo.InvariantCulture))) otherMonths++;
                dateCounted = true;
                for (int index = 0; index < document.Attachments.Count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    EmailAttachment attachment = document.Attachments[index];
                    long length = attachment.Content?.LongLength ?? attachment.ContentSource?.Length ?? attachment.Length;
                    if (length <= 0) { unknownSizes++; continue; }
                    if (length < policy.LargeAttachmentThresholdBytes) continue;
                    largeCount++;
                    var entry = new EmailArchiveLargeAttachment(item.Id, index, Preview(attachment.FileName), length,
                        attachment.MapiAttachMethod == 5 || attachment.EmbeddedDocument != null ||
                        string.Equals(attachment.ContentType, "message/rfc822", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(attachment.ContentType, "message/global", StringComparison.OrdinalIgnoreCase));
                    int insertion = large.FindIndex(existing => existing.DeclaredBytes < length);
                    if (insertion < 0) insertion = large.Count;
                    if (insertion >= policy.MaxLargeAttachments) continue;
                    large.Insert(insertion, entry);
                    if (large.Count > policy.MaxLargeAttachments) large.RemoveAt(large.Count - 1);
                }
                if (item.ContentAvailability.IsHeaderOnly == true || item.ContentAvailability.IsPotentiallyPartial) { partial++; continue; }
                string key = EmailSemanticComparer.CreateFingerprint(document, semanticPolicy, cancellationToken).HexDigest;
                if (!candidates.TryGetValue(key, out var candidate)) candidates.Add(key, candidate = new ArchiveCandidateAccumulator { FirstOrdinal = scanned });
                candidate.Count++;
                if (candidate.Ids.Count < policy.MaxItemsPerDuplicateGroup) candidate.Ids.Add(item.Id);
                long declared = Math.Max(0, document.MessageMetadata.DeclaredSize ?? 0);
                candidate.DeclaredTotal = checked(candidate.DeclaredTotal + declared);
                candidate.LargestDeclared = Math.Max(candidate.LargestDeclared, declared);
            } catch (Exception exception) when (exception is IOException || exception is NotSupportedException || exception is KeyNotFoundException) {
                failed++;
                // A date may already have been counted when semantic matching failed.
                if (!dateCounted) unknownDates++;
                AddDiagnostic(new EmailStoreDiagnostic("EMAIL_ARCHIVE_ANALYSIS_ITEM_FAILED",
                    "The selected item could not be fully analyzed (" + exception.GetType().Name + ").",
                    EmailStoreDiagnosticSeverity.Warning, reference.Id));
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (!string.Equals(fingerprint, GetDurableSourceFingerprint(cancellationToken), StringComparison.Ordinal))
            throw new IOException("The archive source changed during analysis.");
        bool sourceWarnings = false;
        foreach (EmailStoreDiagnostic diagnostic in Diagnostics) {
            sourceWarnings |= diagnostic.Severity != EmailStoreDiagnosticSeverity.Information;
            AddDiagnostic(new EmailStoreDiagnostic(diagnostic.Code, "The source reported " + diagnostic.Code + ".", diagnostic.Severity, diagnostic.Location));
        }
        int groupCount = 0; long estimate = 0; var groups = new List<EmailArchiveDuplicateCandidateGroup>();
        foreach (ArchiveCandidateAccumulator candidate in candidates.Values.OrderBy(candidate => candidate.FirstOrdinal)) {
            if (candidate.Count < 2) continue;
            groupCount++; long bytes = candidate.DeclaredTotal - candidate.LargestDeclared; estimate = checked(estimate + bytes);
            if (groups.Count < policy.MaxDuplicateGroups) groups.Add(new EmailArchiveDuplicateCandidateGroup(candidate.Ids.AsReadOnly(), candidate.Count, bytes));
        }
        return new EmailArchiveAnalysisReport(fingerprint, scanned, projected, failed, partial, scanLimit,
            Buckets(folders), otherFolders, Buckets(months), otherMonths, unknownDates, groups.AsReadOnly(), groupCount,
            estimate, large.AsReadOnly(), largeCount, unknownSizes, diagnostics.AsReadOnly(), diagnosticCount, sourceWarnings);

        bool AddBucket(Dictionary<string, int> buckets, string key) {
            if (buckets.TryGetValue(key, out int count)) { buckets[key] = count + 1; return true; }
            if (buckets.Count == policy.MaxDistributionBuckets) return false;
            buckets.Add(key, 1); return true;
        }
        void AddDiagnostic(EmailStoreDiagnostic diagnostic) {
            diagnosticCount++; if (diagnostics.Count < policy.MaxDiagnostics) diagnostics.Add(diagnostic);
        }
        IReadOnlyList<EmailArchiveDistributionBucket> Buckets(Dictionary<string, int> buckets) => Array.AsReadOnly(buckets
            .OrderBy(pair => pair.Key, StringComparer.Ordinal).Select(pair => new EmailArchiveDistributionBucket(pair.Key, pair.Value)).ToArray());
        string? Preview(string? value) {
            if (value == null || value.Length <= 256) return value;
            int length = char.IsHighSurrogate(value[255]) && char.IsLowSurrogate(value[256]) ? 255 : 256;
            return value.Substring(0, length);
        }
    }

    private sealed class ArchiveCandidateAccumulator {
        internal readonly List<string> Ids = new List<string>();
        internal int Count;
        internal int FirstOrdinal;
        internal long DeclaredTotal, LargestDeclared;
    }
}

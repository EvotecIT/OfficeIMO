using OfficeIMO.Email;
using System.Security.Cryptography;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    private EmailStorePstVerificationReport VerifyPstConversion(string destinationPath,
        PstConversionMappingJournal mappings, EmailStorePstConversionOptions options,
        IList<EmailStoreDiagnostic> diagnostics, string? manifestStagingPath,
        CancellationToken cancellationToken) {
        EmailSemanticComparisonOptions semanticOptions = options.VerificationOptions ??
            CreatePrivateVerificationOptions(options.MaxNestedMessageDepth);
        var issues = new List<EmailStorePstVerificationIssue>();
        bool issuesTruncated = false;
        int matched = 0;
        int mismatched = 0;
        int failed = 0;
        string? manifestPath = null;
        using VerificationManifestWriter? manifest = VerificationManifestWriter.TryCreate(
            manifestStagingPath ?? options.VerificationManifestPath,
            manifestStagingPath == null && options.OverwriteExisting, semanticOptions);
        int destinationEnumerationLimit = mappings.Count == int.MaxValue
            ? int.MaxValue
            : Math.Max(1, mappings.Count + 1);
        using EmailStoreSession destination = EmailStoreSession.Open(destinationPath,
            new EmailStoreReaderOptions(
                maxItemCount: destinationEnumerationLimit,
                retainAttachmentContent: false,
                includeAssociatedItems: true,
                includeOrphanedItems: true), cancellationToken);
        int destinationItemCount = destination.EnumerateItems(new EmailStoreEnumerationOptions(
            includeAssociatedItems: true, includeOrphanedItems: true,
            maxItems: destinationEnumerationLimit), cancellationToken).Count();
        int unmappedDestinationItems = Math.Max(0, destinationItemCount - mappings.Count);
        if (unmappedDestinationItems > 0) {
            failed = unmappedDestinationItems;
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_VERIFY_UNMAPPED_DESTINATION",
                string.Concat(unmappedDestinationItems.ToString(CultureInfo.InvariantCulture),
                    " destination item(s) had no source-provenance mapping."),
                EmailStoreDiagnosticSeverity.Error, destinationPath));
            AddVerificationIssue(issues, ref issuesTruncated, options.MaxVerificationIssues,
                new EmailStorePstVerificationIssue(string.Empty, string.Empty, isAssociated: false,
                    "EMAIL_STORE_PST_VERIFY_UNMAPPED_DESTINATION",
                    Array.Empty<EmailSemanticDifference>()));
        }
        var readOptions = new EmailStoreItemReadOptions(
            EmailStoreItemReadParts.All, preferStreamingAttachmentContent: true);

        foreach (PstConversionItemMap mapping in mappings.ReadAll()) {
            cancellationToken.ThrowIfCancellationRequested();
            string status;
            IReadOnlyList<EmailSemanticDifference> differences = Array.Empty<EmailSemanticDifference>();
            EmailSemanticComparisonReport? comparison = null;
            try {
                EmailStoreItem sourceItem = ReadItem(mapping.Source, readOptions, cancellationToken);
                var destinationReference = new EmailStoreItemReference(
                    mapping.DestinationItemId, mapping.DestinationFolderId,
                    mapping.Source.IsAssociated, isOrphaned: false);
                EmailStoreItem destinationItem = destination.ReadItem(
                    destinationReference, readOptions, cancellationToken);
                comparison = EmailSemanticComparer.Compare(
                    sourceItem.Document, destinationItem.Document, semanticOptions, cancellationToken);
                if (comparison.IsMatch) {
                    matched++;
                    status = "MATCH";
                } else {
                    mismatched++;
                    status = "MISMATCH";
                    differences = comparison.Differences;
                    AddVerificationIssue(issues, ref issuesTruncated, options.MaxVerificationIssues,
                        new EmailStorePstVerificationIssue(mapping.Source.Id,
                            mapping.DestinationItemId, mapping.Source.IsAssociated,
                            "EMAIL_STORE_PST_VERIFY_MISMATCH", differences));
                }
            } catch (Exception exception) when (
                exception is InvalidDataException || exception is IOException ||
                exception is NotSupportedException || exception is KeyNotFoundException ||
                exception is EmailStoreLimitExceededException || exception is EmailLimitExceededException) {
                failed++;
                status = "FAILED";
                AddVerificationIssue(issues, ref issuesTruncated, options.MaxVerificationIssues,
                    new EmailStorePstVerificationIssue(mapping.Source.Id,
                        mapping.DestinationItemId, mapping.Source.IsAssociated,
                        "EMAIL_STORE_PST_VERIFY_ITEM_FAILED", Array.Empty<EmailSemanticDifference>()));
            }
            manifest?.Write(mapping.Ordinal, mapping.Source.IsAssociated,
                mapping.Source.IsOrphaned, status, comparison, differences);
        }

        if (manifest != null) manifestPath = manifest.Complete(
            mappings.Count, matched, mismatched, failed);
        if (mismatched > 0) {
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_VERIFY_MISMATCH",
                string.Concat(mismatched.ToString(CultureInfo.InvariantCulture),
                    " converted item(s) did not match the source semantic projection."),
                EmailStoreDiagnosticSeverity.Error, destinationPath));
        }
        if (failed > 0) {
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_VERIFY_FAILED",
                string.Concat(failed.ToString(CultureInfo.InvariantCulture),
                    " converted item(s) could not be reopened and verified."),
                EmailStoreDiagnosticSeverity.Error, destinationPath));
        }
        return new EmailStorePstVerificationReport(
            checked(mappings.Count + unmappedDestinationItems), matched, mismatched, failed,
            issues.AsReadOnly(), issuesTruncated, manifestPath);
    }

    private static EmailSemanticComparisonOptions CreatePrivateVerificationOptions(int maximumDepth) {
        byte[] key = new byte[32];
        using (RandomNumberGenerator generator = RandomNumberGenerator.Create()) generator.GetBytes(key);
        try {
            return new EmailSemanticComparisonOptions(
                EmailSemanticComparisonProfile.Migration, key,
                maxEmbeddedMessageDepth: maximumDepth);
        } finally {
            Array.Clear(key, 0, key.Length);
        }
    }

    private static void AddVerificationIssue(ICollection<EmailStorePstVerificationIssue> issues,
        ref bool truncated, int maximumIssues, EmailStorePstVerificationIssue issue) {
        if (issues.Count < maximumIssues) issues.Add(issue);
        else truncated = true;
    }

}

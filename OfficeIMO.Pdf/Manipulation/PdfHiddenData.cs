using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;

namespace OfficeIMO.Pdf;

/// <summary>
/// Maps the shared hidden-data contract onto the PDF sanitization engine and PDF provenance. The engine removes the selected
/// categories through one verified full rewrite: it inventories the output again and refuses to return it when a selected item
/// remains, and proves that everything outside the selection was preserved.
/// </summary>
internal static class PdfHiddenData {
    internal const string FormatName = "PDF";


    internal static OfficeHiddenDataReport Inspect(byte[] pdf, PdfLoadOptions? readOptions, CancellationToken cancellationToken) {
        Guard.NotNull(pdf, nameof(pdf));
        var collector = new OfficeHiddenDataCollector(FormatName, null);
        PdfSanitizationReport inventory = InspectInventory(pdf, readOptions, OfficeHiddenDataKind.Default, cancellationToken);
        IReadOnlyList<OfficeProvenanceEvidence> provenance = InspectProvenance(pdf, readOptions, cancellationToken);
        Record(collector, inventory);
        foreach (OfficeProvenanceEvidence evidence in provenance) RecordProvenance(collector, evidence, OfficeHiddenDataOutcome.Found);
        return collector.Build();
    }

    internal static PdfHiddenDataRemoval Remove(
        byte[] pdf,
        OfficeHiddenDataOptions? options,
        PdfLoadOptions? readOptions,
        long? maximumOutputBytes,
        CancellationToken cancellationToken) {
        Guard.NotNull(pdf, nameof(pdf));
        options ??= new OfficeHiddenDataOptions();
        if (maximumOutputBytes <= 0L) throw new ArgumentOutOfRangeException(nameof(maximumOutputBytes), "The maximum output size must be positive.");
        var collector = new OfficeHiddenDataCollector(FormatName, options);
        PdfSanitizationReport inventory = InspectInventory(pdf, readOptions, options.Kinds, cancellationToken);
        IReadOnlyList<OfficeProvenanceEvidence> provenance = InspectProvenance(pdf, readOptions, cancellationToken);

        // Content Credentials sign the file's bytes, so the rewrite below would leave them stale. They are removed first through the
        // provenance engine, which keeps malformed carriers when RequireStructurallyValidCarrier is set.
        byte[] current = pdf;
        PdfLoadOptions? currentReadOptions = readOptions;
        if (provenance.Count > 0) {
            OfficeProvenanceRemovalResult removed = PdfProvenance.Remove(pdf, CreateProvenanceRemoval(options.Provenance, pdf.LongLength, cancellationToken), ProvenanceReadOptions(readOptions));
            if (removed.WasReserialized) {
                current = removed.ToArray();
                currentReadOptions = PdfLoadOptions.WithMinimumInputBytes(readOptions, current.LongLength);
            }
        }

        PdfSanitizationOptions policy = CreatePolicy(options.Kinds);
        policy.MaximumOutputBytes = maximumOutputBytes;
        policy.CancellationToken = cancellationToken;
        PdfSanitizationResult sanitized = PdfSanitizer.Sanitize(current, policy, currentReadOptions);
        if (!sanitized.IsSanitized || !sanitized.PreservationReport.IsPreserved) {
            throw new InvalidOperationException("The rewritten PDF did not pass hidden-data and preservation checks; it was not returned.");
        }

        Record(collector, inventory);
        if (provenance.Count > 0) {
            List<OfficeProvenanceEvidence> remaining = InspectProvenance(sanitized.Bytes, sanitized.ReadOptions, cancellationToken).ToList();
            bool selected = collector.Removes(OfficeHiddenDataKind.Provenance);
            foreach (OfficeProvenanceEvidence evidence in provenance) {
                bool removed = !TakeMatch(remaining, evidence);
                OfficeHiddenDataReason reason = removed
                    ? (!selected || !options.Provenance.RemoveC2paManifests ? OfficeHiddenDataReason.WholeFileCredentialsCannotSurvive : OfficeHiddenDataReason.None)
                    : OfficeHiddenDataReason.ContentCredentialsNoLongerVerify;
                RecordProvenance(collector, evidence, removed ? OfficeHiddenDataOutcome.Removed : OfficeHiddenDataOutcome.Retained, reason);
            }
        }
        OfficeHiddenDataReport report = collector.Build();
        return new PdfHiddenDataRemoval(new OfficeHiddenDataResult(report, sanitized.Bytes), sanitized.Bytes, sanitized.ReadOptions,
            sanitized.PreservationReport);
    }

    /// <summary>The engine policy for a hidden-data selection. Comments exclude rich media, which is active content, and file attachment annotations, which are attachments.</summary>
    internal static PdfSanitizationOptions CreatePolicy(OfficeHiddenDataKind kinds) {
        PdfSanitizationContentKind content = PdfSanitizationContentKind.None;
        if (Selects(kinds, OfficeHiddenDataKind.DocumentProperties)) content |= PdfSanitizationContentKind.UserMetadata;
        if (Selects(kinds, OfficeHiddenDataKind.Comments)) content |= PdfSanitizationContentKind.CommentsAndMarkup;
        if (Selects(kinds, OfficeHiddenDataKind.HiddenContent)) content |= PdfSanitizationContentKind.OptionalContent;
        if (Selects(kinds, OfficeHiddenDataKind.Attachments)) content |= PdfSanitizationContentKind.EmbeddedFiles;
        if (Selects(kinds, OfficeHiddenDataKind.ActiveContent)) content |= PdfSanitizationContentKind.Actions;
        if (Selects(kinds, OfficeHiddenDataKind.Bookmarks)) content |= PdfSanitizationContentKind.Bookmarks;
        return new PdfSanitizationOptions {
            ContentKindsToRemove = content,
            SeparateRichMediaFromComments = true,
            RemoveRichMedia = (content & PdfSanitizationContentKind.Actions) != 0
        };
    }

    private static bool Selects(OfficeHiddenDataKind kinds, OfficeHiddenDataKind kind) => (kinds & kind) == kind;

    /// <summary>Lists every category; what metadata it reports as kept follows <paramref name="removal"/>, the selection being applied or previewed.</summary>
    private static PdfSanitizationReport InspectInventory(byte[] pdf, PdfLoadOptions? readOptions, OfficeHiddenDataKind removal, CancellationToken cancellationToken) {
        PdfSanitizationOptions everything = CreatePolicy(OfficeHiddenDataKind.All);
        everything.KeepsElectronicInvoiceMetadata = !Selects(removal, OfficeHiddenDataKind.Attachments);
        everything.CancellationToken = cancellationToken;
        return PdfSanitizer.Inspect(pdf, everything, readOptions);
    }

    private static void Record(OfficeHiddenDataCollector collector, PdfSanitizationReport inventory) {
        Record(collector, inventory, PdfSanitizationContentKind.UserMetadata, OfficeHiddenDataKind.DocumentProperties);
        Record(collector, inventory, PdfSanitizationContentKind.CommentsAndMarkup, OfficeHiddenDataKind.Comments);
        Record(collector, inventory, PdfSanitizationContentKind.OptionalContent, OfficeHiddenDataKind.HiddenContent);
        Record(collector, inventory, PdfSanitizationContentKind.EmbeddedFiles, OfficeHiddenDataKind.Attachments);
        Record(collector, inventory, PdfSanitizationContentKind.Bookmarks, OfficeHiddenDataKind.Bookmarks);
        foreach (PdfSanitizationFinding finding in inventory.Findings) {
            if (finding.Kind == PdfSanitizationFindingKind.EmbeddedFile) continue;
            RecordEngineItem(collector, OfficeHiddenDataKind.ActiveContent, finding.Path, DescribeActiveContent(finding),
                finding.Kind == PdfSanitizationFindingKind.UnsafeUri ? finding.Detail : null);
        }
    }

    private static void Record(OfficeHiddenDataCollector collector, PdfSanitizationReport inventory, PdfSanitizationContentKind source, OfficeHiddenDataKind kind) {
        foreach (PdfSanitizationItem item in inventory.Items) {
            // Content Credentials stored as an attachment are reported once, as provenance.
            if (item.Kind != source || item.IsProvenanceCarrier) continue;
            if (item.IsRetained) {
                OfficeHiddenDataOutcome outcome = collector.Resolve(kind, removed: false, OfficeHiddenDataReason.NeededByDocument, out OfficeHiddenDataReason reason);
                string? note = reason == OfficeHiddenDataReason.NeededByDocument ? item.RetainedNote : null;
                collector.Record(kind, item.Location, item.Description, item.Value, outcome, note, reason: reason);
            } else {
                RecordEngineItem(collector, kind, item.Location, item.Description, item.Value);
            }
        }
    }

    /// <summary>Records an engine-owned item. A selected item is reported removed because the engine refuses to return output that still contains one.</summary>
    private static void RecordEngineItem(OfficeHiddenDataCollector collector, OfficeHiddenDataKind kind, string location, string description, string? value) {
        OfficeHiddenDataOutcome outcome = collector.Resolve(kind, removed: true, OfficeHiddenDataReason.None, out OfficeHiddenDataReason reason);
        collector.Record(kind, location, description, value, outcome, null, reason: reason);
    }

    private static string DescribeActiveContent(PdfSanitizationFinding finding) {
        switch (finding.Kind) {
            case PdfSanitizationFindingKind.UnsafeUri:
                return "Link or base URI with a scheme that is not allowed";
            case PdfSanitizationFindingKind.RichMedia:
                return "Rich media annotation (" + finding.Detail + ")";
            default:
                if (finding.ActionKind == PdfSanitizationActionKind.JavaScript) {
                    return finding.Path.EndsWith("/JavaScript", StringComparison.Ordinal) ? "Document JavaScript name tree" : "JavaScript action";
                }
                return finding.Detail + " action";
        }
    }

    private static void RecordProvenance(OfficeHiddenDataCollector collector, OfficeProvenanceEvidence evidence, OfficeHiddenDataOutcome outcome,
        OfficeHiddenDataReason reason = OfficeHiddenDataReason.None) =>
        collector.Record(OfficeHiddenDataKind.Provenance, evidence.Location, OfficeHiddenDataAssets.DescribeProvenance(evidence),
            OfficeHiddenDataAssets.ProvenanceValue(evidence), outcome, null, evidence, OfficeHiddenDataScope.Package, reason);

    private static IReadOnlyList<OfficeProvenanceEvidence> InspectProvenance(byte[] pdf, PdfLoadOptions? readOptions, CancellationToken cancellationToken) {
        // The PDF read limits have already admitted the file, so the provenance asset limit must not refuse it again.
        var options = new OfficeProvenanceOptions { CancellationToken = cancellationToken };
        if (pdf.LongLength > options.MaxAssetBytes) options.MaxAssetBytes = pdf.LongLength;
        return PdfProvenance.Inspect(pdf, options, ProvenanceReadOptions(readOptions)).Evidence;
    }

    // Like the sanitization inventory, finding and removing carriers is part of an authorized rewrite, not content extraction for
    // the caller, so the provenance check does not demand the extraction permission that a restricted PDF withholds.
    private static PdfLoadOptions ProvenanceReadOptions(PdfLoadOptions? readOptions) =>
        PdfLoadOptions.WithPermissionPolicy(readOptions, PdfPermissionPolicy.IgnoreRestrictions);

    private static OfficeProvenanceRemovalOptions CreateProvenanceRemoval(OfficeProvenanceRemovalOptions carriers, long length, CancellationToken cancellationToken) {
        var options = new OfficeProvenanceRemovalOptions {
            RemoveC2paManifests = true,
            RemoveExternalC2paReferences = carriers.RemoveExternalC2paReferences,
            RemoveAiSourceMetadata = carriers.RemoveAiSourceMetadata,
            RequireStructurallyValidCarrier = carriers.RequireStructurallyValidCarrier
        };
        options.Limits.MaxAssetBytes = Math.Max(carriers.Limits.MaxAssetBytes, length);
        options.Limits.MaxManifestBytes = carriers.Limits.MaxManifestBytes;
        options.Limits.MaxCarriers = carriers.Limits.MaxCarriers;
        options.Limits.MaxContainerEntries = carriers.Limits.MaxContainerEntries;
        options.Limits.MaxExpandedContainerBytes = carriers.Limits.MaxExpandedContainerBytes;
        options.Limits.CancellationToken = cancellationToken;
        return options;
    }

    private static bool TakeMatch(List<OfficeProvenanceEvidence> remaining, OfficeProvenanceEvidence evidence) {
        int index = remaining.FindIndex(item => item.Carrier == evidence.Carrier && item.Value == evidence.Value &&
            item.PayloadLength == evidence.PayloadLength);
        if (index < 0) return false;
        remaining.RemoveAt(index);
        return true;
    }
}

/// <summary>A verified hidden-data removal: the shared result plus what PDF callers need to keep working with the output.</summary>
internal sealed class PdfHiddenDataRemoval {
    internal PdfHiddenDataRemoval(OfficeHiddenDataResult result, byte[] bytes, PdfLoadOptions readOptions, PdfRewritePreservationReport preservation) {
        Result = result;
        Bytes = bytes;
        ReadOptions = readOptions;
        Preservation = preservation;
    }

    internal OfficeHiddenDataResult Result { get; }

    /// <summary>The cleaned bytes, shared with <see cref="Result"/>; never mutate them.</summary>
    internal byte[] Bytes { get; }

    /// <summary>Read options that admit the cleaned output.</summary>
    internal PdfLoadOptions ReadOptions { get; }

    /// <summary>Proof that everything outside the selection was preserved.</summary>
    internal PdfRewritePreservationReport Preservation { get; }

    /// <summary>Number of findings removed.</summary>
    internal int RemovedCount => Result.Report.Findings.Count(static finding => finding.Outcome == OfficeHiddenDataOutcome.Removed);
}

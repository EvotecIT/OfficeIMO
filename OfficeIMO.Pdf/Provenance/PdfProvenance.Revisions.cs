using OfficeIMO.Provenance;

namespace OfficeIMO.Pdf;

public static partial class PdfProvenance {
    private sealed record HistoricalCarrier(OfficeProvenanceEvidence Evidence, int Spec, string Definition, int Stream, int StreamOffset);
    private readonly record struct CarrierIdentity(int Spec, string Definition, int Stream, int StreamOffset, string Name);
    private readonly record struct AssociationIdentity(int Spec, string Content, bool Associated);
    private readonly record struct CurrentCarrier(int EvidenceIndex, bool IsStructurallyValid);

    // Filespec metadata may be indirect and independently replaced without rewriting the
    // filespec itself. Removal must qualify the entire effective definition in this snapshot.
    private static (string Identity, int Revision) GetCarrierDefinition(PdfReadDocument document, int specNumber,
        OfficeProvenanceOptions options) {
        if (!document.Objects.TryGetValue(specNumber, out var spec)) return (string.Empty, -1);
        var references = CollectReachableObjectNumbers(document.Objects, new PdfReference(specNumber, spec.Generation),
            options.MaxContainerEntries, options.CancellationToken);
        int revision = -1;
        var identities = new List<string>();
        foreach (int number in references.OrderBy(number => number)) {
            options.CancellationToken.ThrowIfCancellationRequested();
            var value = document.Objects[number];
            identities.Add($"{number}:{value.Generation}:{value.SourceOffset}");
            revision = Math.Max(revision, value.SourceRevision);
        }
        return (string.Join(";", identities), revision);
    }

    private static void InspectRevisionCarriers(PdfReadDocument document, OfficeProvenanceOptions options,
        long maximumManifestBytes, List<OfficeProvenanceEvidence> evidence,
        List<ManifestOccurrence> stores, Dictionary<AssociationIdentity, ManifestOccurrence> occurrences, int snapshot,
        Dictionary<CarrierIdentity, int> seen, List<HistoricalCarrier>? historical,
        Dictionary<PdfExtractedAttachment, CurrentCarrier>? currentCarriers = null) {
        var pageTree = CollectPageTreeObjectNumbers(document, options.MaxContainerEntries);
        var associations = CollectAssociationProfile(document, pageTree, options.MaxContainerEntries, out var reachable);
        var attachments = PdfAttachmentExtractor.ExtractAttachments(document, IsCandidate, maximumManifestBytes,
            options.MaxManifestBytes, options.MaxCarriers, options.MaxContainerEntries, requireSuccessfulDecoding: true,
            allowedObjectNumbers: reachable, cancellationToken: options.CancellationToken);
        foreach (var attachment in attachments) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!IsCandidate(attachment)) continue;
            var definition = GetCarrierDefinition(document, attachment.FileSpecObjectNumber, options);
            int streamOffset = document.Objects.TryGetValue(attachment.EmbeddedFileObjectNumber, out var stream) ? stream.SourceOffset : -1;
            // Compressed filespecs share their object-stream offset. Object numbers keep
            // those definitions distinct; direct dictionaries have no stable cross-revision identity.
            var identity = new CarrierIdentity(attachment.FileSpecObjectNumber, definition.Identity,
                attachment.EmbeddedFileObjectNumber, streamOffset, attachment.FileName);
            byte[] manifest = attachment.Bytes;
            if (manifest.LongLength > options.MaxManifestBytes) throw new InvalidDataException("A PDF provenance manifest exceeds the configured manifest limit.");
            bool associated = attachment.Relationship == PdfAssociatedFileRelationship.C2paManifest &&
                string.Equals(attachment.MimeType, C2paMimeType, StringComparison.OrdinalIgnoreCase) &&
                attachment.FileSpecObjectNumber > 0 && HasEmbeddedFileStreamType(document.Objects, attachment) &&
                IsFileSpecificationObject(document.Objects, attachment.FileSpecObjectNumber, pageTree, associations.StructuralObjectNumbers) &&
                HasOnlySelectedEmbeddedFileVariants(document.Objects, attachment) && associations.IsValid(attachment.FileSpecObjectNumber);
            bool valid = associated && OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, options.MaxManifestBytes, options.MaxContainerEntries, out _);
            int evidenceIndex = evidence.Count;
            bool known = attachment.FileSpecObjectNumber > 0 && seen.TryGetValue(identity, out evidenceIndex);
            if (!known) evidenceIndex = evidence.Count;
            if (currentCarriers != null) currentCarriers.Add(attachment, new CurrentCarrier(evidenceIndex, valid));
            // Walk snapshots newest first. A carrier retained in consecutive updates belongs
            // to the earliest of those updates; disappearance and re-association start a new event.
            // Physical rewrites and captions do not create a credential. Track retained
            // associations by owner and logical manifest content; disappearance still ends them.
            var summary = valid ? OfficeC2paManifestStore.TryDescribe(manifest, 0, manifest.Length) : null;
            string content;
            if (summary != null) content = string.Join(";", summary.ManifestIdentities);
            else {
#if NET472 || NETSTANDARD2_0
                using var hash = System.Security.Cryptography.SHA256.Create();
                content = Convert.ToBase64String(hash.ComputeHash(manifest));
#else
                content = Convert.ToBase64String(System.Security.Cryptography.SHA256.HashData(manifest));
#endif
            }
            var key = new AssociationIdentity(attachment.FileSpecObjectNumber, content, associated);
            ManifestOccurrence? occurrence = null;
            bool retained = attachment.FileSpecObjectNumber > 0 && occurrences.TryGetValue(key, out occurrence) && occurrence.LastSnapshot == snapshot + 1;
            if (retained) {
                occurrence!.LastSnapshot = snapshot;
                if (associated && occurrence.Revision >= 0) occurrence.Revision = snapshot;
            }
            if (known && (evidence[evidenceIndex].IsStructurallyValid || !valid)) {
                if (!retained) Observe(evidence[evidenceIndex]);
                continue;
            }
            if (!known && evidence.Count >= options.MaxCarriers) throw new InvalidDataException($"The asset exceeds the configured carrier limit of {options.MaxCarriers}.");
            var carrier = new OfficeProvenanceEvidence(OfficeProvenanceCarrierKind.C2paManifest,
                $"PDF/Filespec[{attachment.FileSpecObjectNumber}]/{attachment.FileName}", valid, manifest.LongLength);
            if (known) {
                evidence[evidenceIndex] = carrier;
                historical?.RemoveAll(item => item.Spec == attachment.FileSpecObjectNumber && item.Definition == definition.Identity &&
                    item.Stream == attachment.EmbeddedFileObjectNumber && item.StreamOffset == streamOffset && item.Evidence.Location == carrier.Location);
            } else {
                if (attachment.FileSpecObjectNumber > 0) seen[identity] = evidence.Count;
                evidence.Add(carrier);
            }
            historical?.Add(new HistoricalCarrier(carrier, attachment.FileSpecObjectNumber, definition.Identity, attachment.EmbeddedFileObjectNumber, streamOffset));
            // Retain candidate ordering even when its carrier or payload is unreadable.
            // Otherwise an older readable credential could be mistaken for the current claim.
            if (!retained) Observe(carrier);

            void Observe(OfficeProvenanceEvidence carrierEvidence) {
                int order = stream != null && stream.SourceRevision >= 0
                    ? associated ? snapshot : Math.Max(stream.SourceRevision, definition.Revision) : -1;
                var candidate = new ManifestOccurrence(carrierEvidence, order, streamOffset,
                    summary) { LastSnapshot = snapshot };
                stores.Add(candidate);
                if (attachment.FileSpecObjectNumber > 0) occurrences[key] = candidate;
            }
        }
    }

    private static long InspectHistoricalCarriers(byte[] pdf, PdfReadDocument current, OfficeProvenanceOptions options,
        PdfLoadOptions readOptions, System.Diagnostics.Stopwatch timer, long maximumManifestBytes,
        List<OfficeProvenanceEvidence> evidence,
        List<ManifestOccurrence> stores, Dictionary<AssociationIdentity, ManifestOccurrence> occurrences, int[] revisionEnds,
        Dictionary<CarrierIdentity, int> seen, List<HistoricalCarrier> historical) {
        long used = current.DecodedStreamBudget.UsedBytes;
        long initial = used;
        for (int index = 0; index < revisionEnds.Length; index++) {
            int end = revisionEnds[index];
            options.CancellationToken.ThrowIfCancellationRequested();
            CheckHistoryBudget(end);
            used += end; // Prefix copies and repeated scans share the cumulative expanded-container budget.
            byte[] prefix = new byte[end];
            for (int copied = 0; copied < end;) {
                options.CancellationToken.ThrowIfCancellationRequested();
                int count = Math.Min(65536, end - copied);
                Buffer.BlockCopy(pdf, copied, prefix, copied, count);
                copied += count;
            }
            long remaining = options.MaxExpandedContainerBytes - used;
            if (remaining <= 0) throw new InvalidDataException("PDF revision inspection exhausted the cumulative expanded-data limit.");
            var revisionOptions = PdfLoadOptions.WithMaximumContainerEntries(readOptions, options.MaxContainerEntries,
                maximumDecodedStreamBytes: Math.Min(readOptions.Limits.MaxDecodedStreamBytes, remaining),
                maximumTotalDecodedStreamBytes: Math.Min(readOptions.Limits.MaxTotalDecodedStreamBytes, remaining),
                maximumTotalAttachmentBytes: Math.Min(maximumManifestBytes, remaining));
            var revision = PdfReadDocument.OpenOwned(prefix, revisionOptions, options.CancellationToken);
            InspectRevisionCarriers(revision, options, maximumManifestBytes, evidence, stores, occurrences,
                revisionEnds.Length - index, seen, historical);
            CheckHistoryBudget(revision.DecodedStreamBudget.UsedBytes);
            used += revision.DecodedStreamBudget.UsedBytes;
        }
        CheckHistoryBudget(0);
        return used - initial;

        void CheckHistoryBudget(long additional) {
            if (additional > options.MaxExpandedContainerBytes - used)
                throw new InvalidDataException("PDF revision inspection exceeds the cumulative expanded-data limit.");
            if (timer.Elapsed > readOptions.Limits.MaxObjectParsingTime)
                throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectParsingTime,
                    (long)readOptions.Limits.MaxObjectParsingTime.TotalMilliseconds, (long)timer.Elapsed.TotalMilliseconds);
        }
    }
}

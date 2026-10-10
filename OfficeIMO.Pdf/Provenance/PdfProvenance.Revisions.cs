using OfficeIMO.Provenance;

namespace OfficeIMO.Pdf;

public static partial class PdfProvenance {
    private sealed record HistoricalCarrier(OfficeProvenanceEvidence Evidence, int Spec, int SpecOffset, int Stream, int StreamOffset);
    private readonly record struct CarrierIdentity(int Spec, int SpecOffset, int Stream, int StreamOffset, string Name);

    private static void InspectRevisionCarriers(PdfReadDocument document, OfficeProvenanceOptions options,
        long maximumManifestBytes, List<OfficeProvenanceEvidence> evidence,
        List<(OfficeProvenanceEvidence Evidence, int Revision, int Stream, OfficeC2paManifestSummary? Summary)> stores,
        Dictionary<CarrierIdentity, int> seen, List<HistoricalCarrier>? historical,
        Dictionary<PdfExtractedAttachment, int>? currentCarriers = null) {
        var pageTree = CollectPageTreeObjectNumbers(document, options.MaxContainerEntries);
        var associations = CollectAssociationProfile(document, pageTree, options.MaxContainerEntries, out var reachable);
        var attachments = PdfAttachmentExtractor.ExtractAttachments(document, IsCandidate, maximumManifestBytes,
            options.MaxManifestBytes, options.MaxCarriers, options.MaxContainerEntries, requireSuccessfulDecoding: true,
            allowedObjectNumbers: reachable, cancellationToken: options.CancellationToken);
        foreach (var attachment in attachments) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!IsCandidate(attachment)) continue;
            int specOffset = document.Objects.TryGetValue(attachment.FileSpecObjectNumber, out var spec) ? spec.SourceOffset : -1;
            int streamOffset = document.Objects.TryGetValue(attachment.EmbeddedFileObjectNumber, out var stream) ? stream.SourceOffset : -1;
            // Compressed filespecs share their object-stream offset. Object numbers keep
            // those definitions distinct; direct dictionaries have no stable cross-revision identity.
            var identity = new CarrierIdentity(attachment.FileSpecObjectNumber, specOffset,
                attachment.EmbeddedFileObjectNumber, streamOffset, attachment.FileName);
            byte[] manifest = attachment.Bytes;
            if (manifest.LongLength > options.MaxManifestBytes) throw new InvalidDataException("A PDF provenance manifest exceeds the configured manifest limit.");
            bool valid = attachment.Relationship == PdfAssociatedFileRelationship.C2paManifest &&
                string.Equals(attachment.MimeType, C2paMimeType, StringComparison.OrdinalIgnoreCase) &&
                attachment.FileSpecObjectNumber > 0 && HasEmbeddedFileStreamType(document.Objects, attachment) &&
                IsFileSpecificationObject(document.Objects, attachment.FileSpecObjectNumber, pageTree, associations.StructuralObjectNumbers) &&
                HasOnlySelectedEmbeddedFileVariants(document.Objects, attachment) && associations.IsValid(attachment.FileSpecObjectNumber) &&
                OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, options.MaxManifestBytes, options.MaxContainerEntries, out _);
            int evidenceIndex = evidence.Count;
            bool known = attachment.FileSpecObjectNumber > 0 && seen.TryGetValue(identity, out evidenceIndex);
            if (!known) evidenceIndex = evidence.Count;
            if (currentCarriers != null) currentCarriers.Add(attachment, evidenceIndex);
            if (known && (evidence[evidenceIndex].IsStructurallyValid || !valid)) continue;
            if (!known && evidence.Count >= options.MaxCarriers) throw new InvalidDataException($"The asset exceeds the configured carrier limit of {options.MaxCarriers}.");
            var carrier = new OfficeProvenanceEvidence(OfficeProvenanceCarrierKind.C2paManifest,
                $"PDF/Filespec[{attachment.FileSpecObjectNumber}]/{attachment.FileName}", valid, manifest.LongLength);
            if (known) {
                evidence[evidenceIndex] = carrier;
                historical?.RemoveAll(item => item.Spec == attachment.FileSpecObjectNumber && item.SpecOffset == specOffset &&
                    item.Stream == attachment.EmbeddedFileObjectNumber && item.StreamOffset == streamOffset && item.Evidence.Location == carrier.Location);
            } else {
                if (attachment.FileSpecObjectNumber > 0) seen[identity] = evidence.Count;
                evidence.Add(carrier);
            }
            historical?.Add(new HistoricalCarrier(carrier, attachment.FileSpecObjectNumber, specOffset, attachment.EmbeddedFileObjectNumber, streamOffset));
            // Retain candidate ordering even when its carrier or payload is unreadable.
            // Otherwise an older readable credential could be mistaken for the current claim.
            stores.Add((carrier, ManifestRevision(document, attachment.EmbeddedFileObjectNumber), streamOffset,
                valid ? OfficeC2paManifestStore.TryDescribe(manifest, 0, manifest.Length) : null));
        }
    }

    private static long InspectHistoricalCarriers(byte[] pdf, PdfReadDocument current, OfficeProvenanceOptions options,
        PdfLoadOptions readOptions, System.Diagnostics.Stopwatch timer, long maximumManifestBytes,
        List<OfficeProvenanceEvidence> evidence,
        List<(OfficeProvenanceEvidence Evidence, int Revision, int Stream, OfficeC2paManifestSummary? Summary)> stores,
        Dictionary<CarrierIdentity, int> seen, List<HistoricalCarrier> historical) {
        long used = current.DecodedStreamBudget.UsedBytes;
        long initial = used;
        foreach (int end in PdfSyntax.GetHistoricalRevisionEnds(pdf, current, options.CancellationToken)) {
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
            InspectRevisionCarriers(revision, options, maximumManifestBytes, evidence, stores, seen, historical);
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

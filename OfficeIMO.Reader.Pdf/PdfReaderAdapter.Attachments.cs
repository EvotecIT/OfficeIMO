using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

internal static partial class PdfReaderAdapter {
    // Embedded files and portfolio items are read through the active Reader's handler set, like ZIP members and
    // email attachments: each rich result is recorded as a nested document and its chunks follow the page chunks.
    private static IEnumerable<ReaderChunk> ReadAttachmentChunks(PdfDocument? pdf, PdfDocumentReadResult document,
        SourceMetadata source, ReaderOptions readerOptions, ReaderPdfOptions pdfOptions, CancellationToken cancellationToken) {
        if (pdf is null || !pdfOptions.IncludeAttachmentContent) yield break;
        IReadOnlyList<PdfAttachmentInfo> attachments = document.Attachments;
        if (attachments.Count == 0) yield break;
        bool portfolio = pdf.Portfolio is not null;
        var attachmentBudget = new AttachmentByteBudget(pdf.ReadOptions.Limits.MaxTotalAttachmentBytes);
        for (int index = 0; index < attachments.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfAttachmentInfo attachment = attachments[index];
            string sourceName = ReaderNestedContent.ResolveSourceName(GetAttachmentFileName(attachment, index), attachment.MimeType);
            // Without a handler the attachment stays a metadata record, as it is for direct adapter calls.
            if (!ReaderNestedContent.CanRead(sourceName)) continue;
            string containerPath = source.Path + "!/attachments/" + index.ToString("D4", CultureInfo.InvariantCulture) + "/" + sourceName;
            OfficeDocumentReadResult? nested = ReadAttachment(pdf, attachment, containerPath, readerOptions, attachmentBudget,
                cancellationToken, out long payloadLength, out string? payloadHash, out string? warning);
            if (nested is null) {
                yield return BuildAttachmentWarningChunk(source, containerPath, index, warning!, readerOptions.ComputeHashes);
                continue;
            }

            nested.Metadata = nested.Metadata
                .Concat(new[] { BuildContainedAttachmentMetadataEntry(attachment, index, portfolio, containerPath, payloadLength) })
                .ToArray();
            ApplyAttachmentSourceIdentity(nested, containerPath, attachment, payloadLength, payloadHash, readerOptions.ComputeHashes);
            string idPrefix = "pdf-attachment-" + index.ToString("D4", CultureInfo.InvariantCulture) + "-";
            foreach (ReaderChunk child in nested.Chunks) {
                cancellationToken.ThrowIfCancellationRequested();
                // Chunks of an attachment's own attachments already carry their deeper path and source identity.
                ReaderChunk chunk = child.CopyForContainer();
                chunk.Id = idPrefix + child.Id;
                chunk.Location = ReaderHierarchicalChunker.CloneLocation(child.Location);
                if (!IsWithinContainer(chunk.Location.Path, containerPath)) chunk.Location.Path = containerPath;
                yield return chunk;
            }
        }
    }

    // The nested source is read under its full container path, so the handler still selects by extension while its
    // own nested results and chunk locations stay rooted at the outermost source.
    private static OfficeDocumentReadResult? ReadAttachment(PdfDocument pdf, PdfAttachmentInfo attachment,
        string containerPath, ReaderOptions readerOptions, AttachmentByteBudget attachmentBudget, CancellationToken cancellationToken,
        out long payloadLength, out string? payloadHash, out string? warning) {
        warning = null;
        payloadLength = 0;
        payloadHash = null;
        // Decoding is bounded before allocation by the remaining nested-input budget and the per-input limit;
        // OfficeIMO.Pdf additionally applies its decoded-stream and attachment limits to every filter chain.
        ReaderOperationBudget? budget = ReaderReadScope.Current?.Budget;
        long? remainingNestedBytes = budget?.RemainingNestedBytes;
        long maximum = Math.Min(remainingNestedBytes ?? long.MaxValue, readerOptions.MaxInputBytes ?? long.MaxValue);
        if (remainingNestedBytes == 0) {
            throw new ReaderResourceLimitException(nameof(ReaderResourceLimits.MaxNestedInputBytes), budget!.NestedInputMaximum!.Value);
        }
        // Each extraction has its own native budget, so the PDF's aggregate attachment limit is carried across them here.
        long remainingAttachmentBytes = attachmentBudget.Remaining(attachment);
        if (remainingAttachmentBytes <= 0) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.AttachmentBytes, attachmentBudget.Maximum, attachmentBudget.Maximum + 1L);
        }
        bool aggregateLimited = remainingAttachmentBytes < maximum;
        maximum = Math.Min(maximum, remainingAttachmentBytes);

        byte[] payload;
        try {
            payload = pdf.Attachments.Extract(attachment, Math.Max(1L, maximum), cancellationToken).PayloadWithoutCopy;
        } catch (PdfReadLimitException error) when (error.Kind == PdfReadLimitKind.AttachmentBytes && error.Limit == maximum) {
            if (aggregateLimited) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.AttachmentBytes, attachmentBudget.Maximum,
                    attachmentBudget.Maximum - remainingAttachmentBytes + error.Actual);
            }
            if (remainingNestedBytes == maximum) {
                throw new ReaderResourceLimitException(nameof(ReaderResourceLimits.MaxNestedInputBytes), budget!.NestedInputMaximum!.Value, error);
            }
            warning = "Skipped PDF attachment because it exceeds MaxInputBytes (" + maximum.ToString(CultureInfo.InvariantCulture) + ").";
            return null;
        } catch (Exception error) when (IsRecoverableAttachmentFailure(error)) {
            warning = "Skipped PDF attachment because its payload could not be decoded: " + error.GetType().Name + ".";
            return null;
        }

        attachmentBudget.Charge(attachment, payload.LongLength);
        payloadLength = payload.LongLength;
        if (readerOptions.ComputeHashes) payloadHash = ComputeSha256Hex(payload);
        try {
            using var content = new MemoryStream(payload, 0, payload.Length, writable: false, publiclyVisible: true);
            return ReaderNestedContent.ReadDocumentInContainer(content, containerPath, containerPath, readerOptions, cancellationToken);
        } catch (Exception error) when (IsRecoverableAttachmentFailure(error)) {
            warning = "Skipped PDF attachment due to a parse error: " + error.GetType().Name + ".";
            return null;
        }
    }

    private static void ApplyAttachmentSourceIdentity(OfficeDocumentReadResult nested, string containerPath, PdfAttachmentInfo attachment,
        long payloadLength, string? payloadHash, bool computeHashes) {
        string sourceId = BuildSourceId(containerPath);
        DateTime? lastWriteUtc = attachment.ModificationDate?.UtcDateTime;
        nested.Source.Path = containerPath;
        nested.Source.SourceId = sourceId;
        nested.Source.SourceHash ??= payloadHash;
        nested.Source.LengthBytes = payloadLength;
        if (lastWriteUtc.HasValue) nested.Source.LastWriteUtc = lastWriteUtc;
        foreach (ReaderChunk chunk in nested.Chunks) {
            // Descendant attachment chunks keep the identity of the bytes they came from.
            if (IsWithinContainer(chunk.Location.Path, containerPath) && !string.Equals(chunk.Location.Path, containerPath, StringComparison.Ordinal)) continue;
            chunk.Location.Path = containerPath;
            chunk.SourceId = sourceId;
            chunk.SourceHash = nested.Source.SourceHash;
            chunk.SourceLengthBytes = payloadLength;
            if (lastWriteUtc.HasValue) chunk.SourceLastWriteUtc = lastWriteUtc;
            chunk.ChunkHash = computeHashes ? ComputeChunkHash(chunk) : null;
        }
    }

    private static bool IsAttachmentChunk(ReaderChunk chunk, SourceMetadata source) =>
        chunk.Location.Path?.StartsWith(source.Path + "!/attachments/", StringComparison.Ordinal) == true;

    private static bool IsWithinContainer(string? path, string containerPath) =>
        path is not null && (string.Equals(path, containerPath, StringComparison.Ordinal) ||
            path.StartsWith(containerPath + "!/", StringComparison.Ordinal));

    // Charges each distinct embedded stream once against PdfReadLimits.MaxTotalAttachmentBytes for the whole document.
    private sealed class AttachmentByteBudget {
        private readonly HashSet<int> _charged = new HashSet<int>();
        private long _used;
        internal AttachmentByteBudget(long maximum) { Maximum = maximum; }
        internal long Maximum { get; }
        internal long Remaining(PdfAttachmentInfo attachment) =>
            IsCharged(attachment) ? long.MaxValue : Maximum - _used;
        internal void Charge(PdfAttachmentInfo attachment, long bytes) {
            if (IsCharged(attachment)) return;
            if (attachment.EmbeddedFileObjectNumber != 0) _charged.Add(attachment.EmbeddedFileObjectNumber);
            _used += bytes;
        }
        private bool IsCharged(PdfAttachmentInfo attachment) =>
            attachment.EmbeddedFileObjectNumber != 0 && _charged.Contains(attachment.EmbeddedFileObjectNumber);
    }

    // Budgets, native quotas and cancellation remain terminal, as they are for ZIP members and email attachments.
    private static bool IsRecoverableAttachmentFailure(Exception error) =>
        error is not OperationCanceledException and not ReaderResourceLimitException and not ReaderAdmissionLimitException
            and not PdfReadLimitException and not OfficeIMO.OfficeInspectionResourceLimitException;

    private static string GetAttachmentFileName(PdfAttachmentInfo attachment, int index) {
        string? name = !string.IsNullOrWhiteSpace(attachment.UnicodeFileName) ? attachment.UnicodeFileName
            : !string.IsNullOrWhiteSpace(attachment.FileName) ? attachment.FileName
            : attachment.Name;
        // File specifications may carry producer paths; only the final segment names the nested source.
        string fileName = (name ?? string.Empty).Replace('\\', '/');
        int separator = fileName.LastIndexOf('/');
        if (separator >= 0) fileName = fileName.Substring(separator + 1);
        fileName = fileName.Trim();
        return fileName.Length == 0 ? "attachment-" + index.ToString("D4", CultureInfo.InvariantCulture) : fileName;
    }

    private static OfficeDocumentMetadataEntry BuildContainedAttachmentMetadataEntry(PdfAttachmentInfo attachment, int index,
        bool portfolio, string containerPath, long decodedSizeBytes) {
        OfficeDocumentMetadataEntry entry = BuildAttachmentMetadataEntry(attachment, index);
        var attributes = entry.Attributes.ToDictionary(item => item.Key, item => item.Value, StringComparer.Ordinal);
        attributes["containerPath"] = containerPath;
        // A catalog /Collection turns the embedded-files name tree into portfolio items; annotation and
        // associated files remain attachments of their page or structure owner.
        attributes["containerRelationship"] = portfolio && string.Equals(attachment.Source, "Names/EmbeddedFiles", StringComparison.Ordinal)
                ? "portfolio-item"
                : "attachment";
        attributes["decodedSizeBytes"] = decodedSizeBytes.ToString(CultureInfo.InvariantCulture);
        entry.Id = "pdf-container-attachment";
        entry.Category = "pdf.container.attachment";
        entry.Attributes = attributes;
        return entry;
    }

    private static ReaderChunk BuildAttachmentWarningChunk(SourceMetadata source, string containerPath, int index, string warning,
        bool computeHashes) => EnrichChunk(new ReaderChunk {
            Id = "pdf-attachment-" + index.ToString("D4", CultureInfo.InvariantCulture) + "-warning",
            Kind = ReaderInputKind.Pdf,
            Location = new ReaderLocation { Path = containerPath, SourceBlockIndex = index, SourceBlockKind = "attachment" },
            Text = warning,
            Warnings = new[] { warning }
        }, source, computeHashes);
}

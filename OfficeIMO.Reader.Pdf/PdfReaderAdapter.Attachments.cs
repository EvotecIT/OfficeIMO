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
        for (int index = 0; index < attachments.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfAttachmentInfo attachment = attachments[index];
            string sourceName = ReaderNestedContent.ResolveSourceName(GetAttachmentFileName(attachment, index), attachment.MimeType);
            // Without a handler the attachment stays a metadata record, as it is for direct adapter calls.
            if (!ReaderNestedContent.CanRead(sourceName)) continue;
            string containerPath = source.Path + "!/attachments/" + index.ToString("D4", CultureInfo.InvariantCulture) + "/" + sourceName;
            OfficeDocumentReadResult? nested = ReadAttachment(pdf, attachment, sourceName, containerPath, readerOptions,
                cancellationToken, out long payloadLength, out string? warning);
            if (nested is null) {
                yield return BuildAttachmentWarningChunk(source, containerPath, index, warning!, readerOptions.ComputeHashes);
                continue;
            }

            nested.Metadata = nested.Metadata
                .Concat(new[] { BuildContainedAttachmentMetadataEntry(attachment, index, portfolio, containerPath, payloadLength) })
                .ToArray();
            string sourceId = BuildSourceId(containerPath);
            DocumentReaderEngine.ApplyExternalSourceMetadata(nested, sourceId,
                attachment.ModificationDate?.UtcDateTime, payloadLength, readerOptions.ComputeHashes);
            string idPrefix = "pdf-attachment-" + index.ToString("D4", CultureInfo.InvariantCulture) + "-";
            foreach (ReaderChunk child in nested.Chunks) {
                cancellationToken.ThrowIfCancellationRequested();
                ReaderChunk chunk = child.CopyForContainer();
                chunk.Id = idPrefix + child.Id;
                chunk.Location = ReaderHierarchicalChunker.CloneLocation(child.Location);
                chunk.Location.Path = containerPath;
                chunk.SourceId = sourceId;
                chunk.SourceHash = source.SourceHash;
                chunk.SourceLastWriteUtc = source.LastWriteUtc;
                chunk.SourceLengthBytes = payloadLength;
                chunk.ChunkHash = null;
                yield return EnrichChunk(chunk, source, readerOptions.ComputeHashes);
            }
        }
    }

    private static OfficeDocumentReadResult? ReadAttachment(PdfDocument pdf, PdfAttachmentInfo attachment, string sourceName,
        string containerPath, ReaderOptions readerOptions, CancellationToken cancellationToken, out long payloadLength, out string? warning) {
        warning = null;
        payloadLength = 0;
        // Decoding is bounded before allocation by the remaining nested-input budget and the per-input limit;
        // OfficeIMO.Pdf additionally applies its decoded-stream and attachment limits to every filter chain.
        ReaderOperationBudget? budget = ReaderReadScope.Current?.Budget;
        long? remainingNestedBytes = budget?.RemainingNestedBytes;
        long maximum = Math.Min(remainingNestedBytes ?? long.MaxValue, readerOptions.MaxInputBytes ?? long.MaxValue);
        if (remainingNestedBytes == 0) {
            throw new ReaderResourceLimitException(nameof(ReaderResourceLimits.MaxNestedInputBytes), budget!.NestedInputMaximum!.Value);
        }

        byte[] payload;
        try {
            payload = pdf.Attachments.Extract(attachment, Math.Max(1L, maximum), cancellationToken).PayloadWithoutCopy;
        } catch (PdfReadLimitException error) when (error.Kind == PdfReadLimitKind.AttachmentBytes && error.Limit == maximum) {
            if (remainingNestedBytes == maximum) {
                throw new ReaderResourceLimitException(nameof(ReaderResourceLimits.MaxNestedInputBytes), budget!.NestedInputMaximum!.Value, error);
            }
            warning = "Skipped PDF attachment because it exceeds MaxInputBytes (" + maximum.ToString(CultureInfo.InvariantCulture) + ").";
            return null;
        } catch (Exception error) when (IsRecoverableAttachmentFailure(error)) {
            warning = "Skipped PDF attachment because its payload could not be decoded: " + error.GetType().Name + ".";
            return null;
        }

        payloadLength = payload.LongLength;
        try {
            using var content = new MemoryStream(payload, 0, payload.Length, writable: false, publiclyVisible: true);
            return ReaderNestedContent.ReadDocumentInContainer(content, sourceName, containerPath, readerOptions, cancellationToken);
        } catch (Exception error) when (IsRecoverableAttachmentFailure(error)) {
            warning = "Skipped PDF attachment due to a parse error: " + error.GetType().Name + ".";
            return null;
        }
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

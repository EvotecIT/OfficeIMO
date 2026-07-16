using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;

namespace OfficeIMO.Word.GoogleDocs {
    public sealed partial class GoogleDocsExporter : IGoogleDocsExporter {
        private static async Task ApplyDocumentContentAsync(
            GoogleWorkspaceHttpTransport transport,
            string accessToken,
            string documentId,
            GoogleDocsBatch batch,
            GoogleDocsSaveOptions options,
            GoogleDriveClient driveClient,
            CancellationToken cancellationToken) {
            var leases = new List<GoogleDriveTemporaryContentLease>();
            try {
                IReadOnlyDictionary<GoogleDocsInlineImage, string> imageUris = options.InlineImageMode == GoogleDocsInlineImageMode.TemporaryPublicDriveLease
                    ? await UploadInlineImagesAsync(driveClient, batch, leases, cancellationToken).ConfigureAwait(false)
                    : new Dictionary<GoogleDocsInlineImage, string>();

                string? originalTargetTabId = batch.TargetTabId;
                try {
                    if (batch.TargetTabIds.Count == 0) {
                        await ApplyDocumentContentCoreAsync(
                            transport,
                            accessToken,
                            documentId,
                            batch,
                            imageUris,
                            cancellationToken).ConfigureAwait(false);
                    } else {
                        foreach (string targetTabId in batch.TargetTabIds) {
                            batch.TargetTabId = targetTabId;
                            await ApplyDocumentContentCoreAsync(
                                transport,
                                accessToken,
                                documentId,
                                batch,
                                imageUris,
                                cancellationToken).ConfigureAwait(false);
                        }
                    }
                } finally {
                    batch.TargetTabId = originalTargetTabId;
                }
            } finally {
                await CleanupTemporaryInlineImagesAsync(
                    leases,
                    batch.Report,
                    CancellationToken.None).ConfigureAwait(false);
            }
        }

        private static async Task ApplyDocumentContentCoreAsync(
            GoogleWorkspaceHttpTransport transport,
            string accessToken,
            string documentId,
            GoogleDocsBatch batch,
            IReadOnlyDictionary<GoogleDocsInlineImage, string> imageUris,
            CancellationToken cancellationToken) {
            var preparedInitialBatch = GoogleDocsApiPayloadBuilder.BuildPreparedInitialBatchUpdate(batch, imageUris);
            GoogleDocsApiBatchUpdateResponse? initialResponse = null;
            var initialPayload = preparedInitialBatch.Payload;
            if (initialPayload.Requests.Count > 0) {
                initialResponse = await SendBatchUpdateAsync(transport, accessToken, documentId, batch, initialPayload, cancellationToken).ConfigureAwait(false);
            }

            if (preparedInitialBatch.Footnotes.Count > 0 && initialResponse != null) {
                await ApplyFootnotesAsync(
                    transport,
                    accessToken,
                    documentId,
                    batch,
                    preparedInitialBatch.Footnotes,
                    initialResponse,
                    imageUris,
                    cancellationToken).ConfigureAwait(false);
            }

            bool needsDocumentState = batch.Requests.OfType<GoogleDocsInsertTableRequest>().Any()
                || batch.Segments.Any(segment =>
                    string.Equals(segment.Variant, "default", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(segment.Variant, "first", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(segment.Variant, "even", StringComparison.OrdinalIgnoreCase));

            if (!needsDocumentState) {
                return;
            }

            var documentState = await GetDocumentAsync(transport, accessToken, documentId, batch, cancellationToken).ConfigureAwait(false);

            await ApplyHeaderFooterSegmentsAsync(
                transport,
                accessToken,
                documentId,
                batch,
                imageUris,
                documentState,
                cancellationToken).ConfigureAwait(false);

            if (batch.Requests.OfType<GoogleDocsInsertTableRequest>().Any()) {
                var preparedTableBatch = GoogleDocsApiPayloadBuilder.BuildPreparedTableContentBatchUpdate(batch, documentState, imageUris);
                GoogleDocsApiBatchUpdateResponse? tableContentResponse = null;
                var tablePayload = preparedTableBatch.Payload;
                if (tablePayload.Requests.Count > 0) {
                    tableContentResponse = await SendBatchUpdateAsync(transport, accessToken, documentId, batch, tablePayload, cancellationToken).ConfigureAwait(false);
                }

                if (preparedTableBatch.Footnotes.Count > 0 && tableContentResponse != null) {
                    await ApplyFootnotesAsync(
                        transport,
                        accessToken,
                        documentId,
                        batch,
                        preparedTableBatch.Footnotes,
                        tableContentResponse,
                        imageUris,
                        cancellationToken).ConfigureAwait(false);
                }

                if (tablePayload.Requests.Count > 0) {
                    documentState = await GetDocumentAsync(
                        transport,
                        accessToken,
                        documentId,
                        batch,
                        cancellationToken).ConfigureAwait(false);
                }

                var mergePayload = GoogleDocsApiPayloadBuilder.BuildTableMergeBatchUpdatePayload(batch, documentState);
                if (mergePayload.Requests.Count > 0) {
                    await SendBatchUpdateAsync(transport, accessToken, documentId, batch, mergePayload, cancellationToken).ConfigureAwait(false);
                }

                var tableStylePayload = GoogleDocsApiPayloadBuilder.BuildTableStyleBatchUpdatePayload(batch, documentState);
                if (tableStylePayload.Requests.Count > 0) {
                    await SendBatchUpdateAsync(transport, accessToken, documentId, batch, tableStylePayload, cancellationToken).ConfigureAwait(false);
                }
            }
        }


        private static async Task ApplyFootnotesAsync(
            GoogleWorkspaceHttpTransport transport,
            string accessToken,
            string documentId,
            GoogleDocsBatch batch,
            IReadOnlyList<GoogleDocsFootnote> footnotes,
            GoogleDocsApiBatchUpdateResponse initialResponse,
            IReadOnlyDictionary<GoogleDocsInlineImage, string> imageUris,
            CancellationToken cancellationToken) {
            var footnoteReplies = initialResponse.Replies
                .Where(reply => reply.CreateFootnote?.FootnoteId != null)
                .Select(reply => reply.CreateFootnote!.FootnoteId!)
                .ToList();

            if (footnoteReplies.Count != footnotes.Count) {
                batch.Report.Add(
                    TranslationSeverity.Warning,
                    "Footnotes",
                    $"Expected {footnotes.Count} Google Docs footnote replies after creation, but the API returned {footnoteReplies.Count}. Footnote content replay may be incomplete.");
            }

            for (int index = 0; index < Math.Min(footnotes.Count, footnoteReplies.Count); index++) {
                var footnotePayload = GoogleDocsApiPayloadBuilder.BuildFootnoteBatchUpdatePayload(
                    footnotes[index],
                    batch.Report,
                    footnoteReplies[index],
                    imageUris);
                if (footnotePayload.Requests.Count == 0) {
                    continue;
                }

                await SendBatchUpdateAsync(transport, accessToken, documentId, batch, footnotePayload, cancellationToken).ConfigureAwait(false);
            }
        }

    }
}

using OfficeIMO.GoogleWorkspace;
using System.Net.Http.Headers;
using System.IO;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Word.GoogleDocs {
    public sealed partial class GoogleDocsExporter : IGoogleDocsExporter {
        private static async Task ApplyDocumentContentAsync(
            HttpClient client,
            string accessToken,
            string documentId,
            GoogleDocsBatch batch,
            GoogleWorkspaceRetryOptions retryOptions,
            CancellationToken cancellationToken) {
            var imageUris = await UploadInlineImagesAsync(
                client,
                accessToken,
                batch,
                retryOptions,
                cancellationToken).ConfigureAwait(false);

            var preparedInitialBatch = GoogleDocsApiPayloadBuilder.BuildPreparedInitialBatchUpdate(batch, imageUris);
            GoogleDocsApiBatchUpdateResponse? initialResponse = null;
            var initialPayload = preparedInitialBatch.Payload;
            if (initialPayload.Requests.Count > 0) {
                initialResponse = await SendAsync<GoogleDocsApiBatchUpdateResponse>(
                    client,
                    accessToken,
                    HttpMethod.Post,
                    $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                    initialPayload,
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);
            }

            if (preparedInitialBatch.Footnotes.Count > 0 && initialResponse != null) {
                await ApplyFootnotesAsync(
                    client,
                    accessToken,
                    documentId,
                    batch,
                    preparedInitialBatch.Footnotes,
                    initialResponse,
                    imageUris,
                    retryOptions,
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

            var documentState = await SendAsync<GoogleDocsApiDocumentResponse>(
                client,
                accessToken,
                HttpMethod.Get,
                $"https://docs.googleapis.com/v1/documents/{documentId}",
                null,
                retryOptions,
                batch.Report,
                cancellationToken).ConfigureAwait(false);

            await ApplyHeaderFooterSegmentsAsync(
                client,
                accessToken,
                documentId,
                batch,
                imageUris,
                documentState,
                retryOptions,
                cancellationToken).ConfigureAwait(false);

            if (batch.Requests.OfType<GoogleDocsInsertTableRequest>().Any()) {
                var preparedTableBatch = GoogleDocsApiPayloadBuilder.BuildPreparedTableContentBatchUpdate(batch, documentState, imageUris);
                GoogleDocsApiBatchUpdateResponse? tableContentResponse = null;
                var tablePayload = preparedTableBatch.Payload;
                if (tablePayload.Requests.Count > 0) {
                    tableContentResponse = await SendAsync<GoogleDocsApiBatchUpdateResponse>(
                        client,
                        accessToken,
                        HttpMethod.Post,
                        $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                        tablePayload,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);
                }

                if (preparedTableBatch.Footnotes.Count > 0 && tableContentResponse != null) {
                    await ApplyFootnotesAsync(
                        client,
                        accessToken,
                        documentId,
                        batch,
                        preparedTableBatch.Footnotes,
                        tableContentResponse,
                        imageUris,
                        retryOptions,
                        cancellationToken).ConfigureAwait(false);
                }

                var mergePayload = GoogleDocsApiPayloadBuilder.BuildTableMergeBatchUpdatePayload(batch, documentState);
                if (mergePayload.Requests.Count > 0) {
                    await SendAsync<object>(
                        client,
                        accessToken,
                        HttpMethod.Post,
                        $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                        mergePayload,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);
                }

                var tableStylePayload = GoogleDocsApiPayloadBuilder.BuildTableStyleBatchUpdatePayload(batch, documentState);
                if (tableStylePayload.Requests.Count > 0) {
                    await SendAsync<object>(
                        client,
                        accessToken,
                        HttpMethod.Post,
                        $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                        tableStylePayload,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);
                }
            }
        }


        private static async Task ApplyFootnotesAsync(
            HttpClient client,
            string accessToken,
            string documentId,
            GoogleDocsBatch batch,
            IReadOnlyList<GoogleDocsFootnote> footnotes,
            GoogleDocsApiBatchUpdateResponse initialResponse,
            IReadOnlyDictionary<GoogleDocsInlineImage, string> imageUris,
            GoogleWorkspaceRetryOptions retryOptions,
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

                await SendAsync<object>(
                    client,
                    accessToken,
                    HttpMethod.Post,
                    $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                    footnotePayload,
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);
            }
        }

    }
}

using OfficeIMO.GoogleWorkspace;
using System.Net.Http.Headers;
using System.IO;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Word.GoogleDocs {
    public sealed partial class GoogleDocsExporter : IGoogleDocsExporter {
        private static async Task ApplyHeaderFooterSegmentsAsync(
            HttpClient client,
            string accessToken,
            string documentId,
            GoogleDocsBatch batch,
            IReadOnlyDictionary<GoogleDocsInlineImage, string> imageUris,
            GoogleDocsApiDocumentResponse documentState,
            GoogleWorkspaceRetryOptions retryOptions,
            CancellationToken cancellationToken) {
            var executableSegments = batch.Segments
                .Where(segment =>
                    string.Equals(segment.Variant, "default", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(segment.Variant, "first", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(segment.Variant, "even", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (executableSegments.Count == 0) {
                return;
            }

            var sectionBreakIndexes = EnumerateSectionBreakIndexes(documentState).ToList();
            foreach (var segment in executableSegments) {
                string? sectionBreakLocation = null;
                if (segment.SectionIndex > 0) {
                    if (sectionBreakIndexes.Count < segment.SectionIndex) {
                        batch.Report.Add(
                            TranslationSeverity.Warning,
                            "HeadersAndFooters",
                            $"Could not resolve the Google Docs section break location for section {segment.SectionIndex + 1}, so its {segment.Variant} {segment.Kind} was skipped.");
                        continue;
                    }

                    sectionBreakLocation = sectionBreakIndexes[segment.SectionIndex - 1].ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                string? segmentId;
                if (string.Equals(segment.Kind, "header", StringComparison.OrdinalIgnoreCase)) {
                    segmentId = await CreateHeaderAsync(client, accessToken, documentId, sectionBreakLocation, segment.Variant, retryOptions, batch.Report, cancellationToken).ConfigureAwait(false);
                } else {
                    segmentId = await CreateFooterAsync(client, accessToken, documentId, sectionBreakLocation, segment.Variant, retryOptions, batch.Report, cancellationToken).ConfigureAwait(false);
                }

                if (string.IsNullOrWhiteSpace(segmentId)) {
                    batch.Report.Add(
                        TranslationSeverity.Warning,
                        "HeadersAndFooters",
                        $"Google Docs did not return a segment id for section {segment.SectionIndex + 1} {segment.Variant} {segment.Kind}, so its content was skipped.");
                    continue;
                }

                var segmentPayload = GoogleDocsApiPayloadBuilder.BuildSegmentBatchUpdatePayload(segment, batch.Report, segmentId!, imageUris);
                if (segmentPayload.Requests.Count > 0) {
                    await SendAsync<object>(
                        client,
                        accessToken,
                        HttpMethod.Post,
                        $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                        segmentPayload,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);
                }

                if (!segment.Requests.OfType<GoogleDocsInsertTableRequest>().Any()) {
                    continue;
                }

                var segmentDocumentState = await SendAsync<GoogleDocsApiDocumentResponse>(
                    client,
                    accessToken,
                    HttpMethod.Get,
                    $"https://docs.googleapis.com/v1/documents/{documentId}",
                    null,
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);

                var segmentTablePayload = GoogleDocsApiPayloadBuilder.BuildSegmentTableContentBatchUpdatePayload(
                    segment,
                    segmentDocumentState,
                    batch.Report,
                    segmentId!,
                    imageUris);
                if (segmentTablePayload.Requests.Count == 0) {
                    continue;
                }

                await SendAsync<object>(
                    client,
                    accessToken,
                    HttpMethod.Post,
                    $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                    segmentTablePayload,
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);

                var segmentMergePayload = GoogleDocsApiPayloadBuilder.BuildSegmentTableMergeBatchUpdatePayload(
                    segment,
                    segmentDocumentState,
                    batch.Report,
                    segmentId!);
                if (segmentMergePayload.Requests.Count > 0) {
                    await SendAsync<object>(
                        client,
                        accessToken,
                        HttpMethod.Post,
                        $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                        segmentMergePayload,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);
                }

                var segmentTableStylePayload = GoogleDocsApiPayloadBuilder.BuildSegmentTableStyleBatchUpdatePayload(
                    segment,
                    segmentDocumentState,
                    batch.Report,
                    segmentId!);
                if (segmentTableStylePayload.Requests.Count == 0) {
                    continue;
                }

                await SendAsync<object>(
                    client,
                    accessToken,
                    HttpMethod.Post,
                    $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                    segmentTableStylePayload,
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);
            }
        }

        private static IEnumerable<int> EnumerateSectionBreakIndexes(GoogleDocsApiDocumentResponse documentState) {
            var content = documentState.Body?.Content;
            if (content == null) {
                yield break;
            }

            foreach (var element in content) {
                if (element.SectionBreak != null && element.StartIndex.HasValue) {
                    yield return element.StartIndex.Value;
                }
            }
        }

        private static async Task<string?> CreateHeaderAsync(
            HttpClient client,
            string accessToken,
            string documentId,
            string? sectionBreakLocation,
            string variant,
            GoogleWorkspaceRetryOptions retryOptions,
            TranslationReport report,
            CancellationToken cancellationToken) {
            var payload = new GoogleDocsApiBatchUpdatePayload();
            payload.Requests.Add(new GoogleDocsApiRequestPayload {
                CreateHeader = new GoogleDocsApiCreateHeaderRequestPayload {
                    Type = ResolveHeaderFooterType(variant),
                    SectionBreakLocation = string.IsNullOrWhiteSpace(sectionBreakLocation)
                        ? null
                        : new GoogleDocsApiLocationPayload { Index = int.Parse(sectionBreakLocation, System.Globalization.CultureInfo.InvariantCulture) }
                }
            });

            var response = await SendAsync<GoogleDocsApiBatchUpdateResponse>(
                client,
                accessToken,
                HttpMethod.Post,
                $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                payload,
                retryOptions,
                report,
                cancellationToken).ConfigureAwait(false);

            return response.Replies.FirstOrDefault()?.CreateHeader?.HeaderId;
        }

        private static async Task<string?> CreateFooterAsync(
            HttpClient client,
            string accessToken,
            string documentId,
            string? sectionBreakLocation,
            string variant,
            GoogleWorkspaceRetryOptions retryOptions,
            TranslationReport report,
            CancellationToken cancellationToken) {
            var payload = new GoogleDocsApiBatchUpdatePayload();
            payload.Requests.Add(new GoogleDocsApiRequestPayload {
                CreateFooter = new GoogleDocsApiCreateFooterRequestPayload {
                    Type = ResolveHeaderFooterType(variant),
                    SectionBreakLocation = string.IsNullOrWhiteSpace(sectionBreakLocation)
                        ? null
                        : new GoogleDocsApiLocationPayload { Index = int.Parse(sectionBreakLocation, System.Globalization.CultureInfo.InvariantCulture) }
                }
            });

            var response = await SendAsync<GoogleDocsApiBatchUpdateResponse>(
                client,
                accessToken,
                HttpMethod.Post,
                $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                payload,
                retryOptions,
                report,
                cancellationToken).ConfigureAwait(false);

            return response.Replies.FirstOrDefault()?.CreateFooter?.FooterId;
        }

        private static string ResolveHeaderFooterType(string variant) {
            if (string.Equals(variant, "first", StringComparison.OrdinalIgnoreCase)) {
                return "FIRST_PAGE";
            }

            if (string.Equals(variant, "even", StringComparison.OrdinalIgnoreCase)) {
                return "EVEN_PAGE";
            }

            return "DEFAULT";
        }

    }
}

using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Word.GoogleDocs {
    public sealed partial class GoogleDocsExporter : IGoogleDocsExporter {
        private static async Task ApplyHeaderFooterSegmentsAsync(
            GoogleWorkspaceHttpTransport transport,
            string accessToken,
            string documentId,
            GoogleDocsBatch batch,
            IReadOnlyDictionary<GoogleDocsInlineImage, string> imageUris,
            GoogleDocsApiDocumentResponse documentState,
            CancellationToken cancellationToken) {
            var executableSegments = batch.Segments
                .Where(segment => string.Equals(segment.Variant, "default", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (executableSegments.Count == 0) {
                return;
            }

            var sectionBreakIndexes = EnumerateSectionBreakIndexes(documentState).ToList();
            bool selectedTabHasInitialSectionAnchor = !string.IsNullOrWhiteSpace(batch.TargetTabId)
                && sectionBreakIndexes.Count > 0
                && sectionBreakIndexes[0] == 0;
            foreach (var segment in executableSegments) {
                GoogleDocsApiLocationPayload? sectionBreakLocation = null;
                if (segment.SectionIndex > 0 || !string.IsNullOrWhiteSpace(batch.TargetTabId)) {
                    int sectionBreakListIndex = segment.SectionIndex - 1 + (selectedTabHasInitialSectionAnchor ? 1 : 0);
                    if (sectionBreakListIndex < 0 || sectionBreakIndexes.Count <= sectionBreakListIndex) {
                        batch.Report.Add(
                            TranslationSeverity.Warning,
                            "HeadersAndFooters",
                            $"Could not resolve the Google Docs section break location for section {segment.SectionIndex + 1}, so its {segment.Variant} {segment.Kind} was skipped.");
                        continue;
                    }

                    sectionBreakLocation = new GoogleDocsApiLocationPayload {
                        Index = sectionBreakIndexes[sectionBreakListIndex]
                    };
                }

                string? segmentId;
                if (string.Equals(segment.Kind, "header", StringComparison.OrdinalIgnoreCase)) {
                    segmentId = await CreateHeaderAsync(transport, accessToken, documentId, sectionBreakLocation, batch, cancellationToken).ConfigureAwait(false);
                } else {
                    segmentId = await CreateFooterAsync(transport, accessToken, documentId, sectionBreakLocation, batch, cancellationToken).ConfigureAwait(false);
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
                    await SendBatchUpdateAsync(transport, accessToken, documentId, batch, segmentPayload, cancellationToken).ConfigureAwait(false);
                }

                if (!segment.Requests.OfType<GoogleDocsInsertTableRequest>().Any()) {
                    continue;
                }

                var segmentDocumentState = await GetDocumentAsync(transport, accessToken, documentId, batch, cancellationToken).ConfigureAwait(false);

                var segmentTablePayload = GoogleDocsApiPayloadBuilder.BuildSegmentTableContentBatchUpdatePayload(
                    segment,
                    segmentDocumentState,
                    batch.Report,
                    segmentId!,
                    imageUris);
                if (segmentTablePayload.Requests.Count == 0) {
                    continue;
                }

                await SendBatchUpdateAsync(transport, accessToken, documentId, batch, segmentTablePayload, cancellationToken).ConfigureAwait(false);

                segmentDocumentState = await GetDocumentAsync(
                    transport,
                    accessToken,
                    documentId,
                    batch,
                    cancellationToken).ConfigureAwait(false);

                var segmentMergePayload = GoogleDocsApiPayloadBuilder.BuildSegmentTableMergeBatchUpdatePayload(
                    segment,
                    segmentDocumentState,
                    batch.Report,
                    segmentId!);
                if (segmentMergePayload.Requests.Count > 0) {
                    await SendBatchUpdateAsync(transport, accessToken, documentId, batch, segmentMergePayload, cancellationToken).ConfigureAwait(false);
                }

                var segmentTableStylePayload = GoogleDocsApiPayloadBuilder.BuildSegmentTableStyleBatchUpdatePayload(
                    segment,
                    segmentDocumentState,
                    batch.Report,
                    segmentId!);
                if (segmentTableStylePayload.Requests.Count == 0) {
                    continue;
                }

                await SendBatchUpdateAsync(transport, accessToken, documentId, batch, segmentTableStylePayload, cancellationToken).ConfigureAwait(false);
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
            GoogleWorkspaceHttpTransport transport,
            string accessToken,
            string documentId,
            GoogleDocsApiLocationPayload? sectionBreakLocation,
            GoogleDocsBatch batch,
            CancellationToken cancellationToken) {
            var payload = new GoogleDocsApiBatchUpdatePayload();
            payload.Requests.Add(new GoogleDocsApiRequestPayload {
                CreateHeader = new GoogleDocsApiCreateHeaderRequestPayload {
                    Type = "DEFAULT",
                    SectionBreakLocation = sectionBreakLocation
                }
            });

            var response = await SendBatchUpdateAsync(transport, accessToken, documentId, batch, payload, cancellationToken).ConfigureAwait(false);

            return response.Replies.FirstOrDefault()?.CreateHeader?.HeaderId;
        }

        private static async Task<string?> CreateFooterAsync(
            GoogleWorkspaceHttpTransport transport,
            string accessToken,
            string documentId,
            GoogleDocsApiLocationPayload? sectionBreakLocation,
            GoogleDocsBatch batch,
            CancellationToken cancellationToken) {
            var payload = new GoogleDocsApiBatchUpdatePayload();
            payload.Requests.Add(new GoogleDocsApiRequestPayload {
                CreateFooter = new GoogleDocsApiCreateFooterRequestPayload {
                    Type = "DEFAULT",
                    SectionBreakLocation = sectionBreakLocation
                }
            });

            var response = await SendBatchUpdateAsync(transport, accessToken, documentId, batch, payload, cancellationToken).ConfigureAwait(false);

            return response.Replies.FirstOrDefault()?.CreateFooter?.FooterId;
        }
    }
}

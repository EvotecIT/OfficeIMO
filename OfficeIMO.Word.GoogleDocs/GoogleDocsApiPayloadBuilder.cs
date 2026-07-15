using OfficeIMO.GoogleWorkspace;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.GoogleDocs {
    internal static partial class GoogleDocsApiPayloadBuilder {
        internal static GoogleDocsApiCreateDocumentPayload BuildCreateDocumentPayload(GoogleDocsBatch batch) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));
            return new GoogleDocsApiCreateDocumentPayload {
                Title = batch.Title,
            };
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildInitialBatchUpdatePayload(GoogleDocsBatch batch) {
            return BuildPreparedInitialBatchUpdate(batch, null).Payload;
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildInitialBatchUpdatePayload(
            GoogleDocsBatch batch,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris) {
            return BuildPreparedInitialBatchUpdate(batch, imageUris).Payload;
        }

        internal static PreparedInitialBatchUpdate BuildPreparedInitialBatchUpdate(
            GoogleDocsBatch batch,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));

            var prepared = new PreparedInitialBatchUpdate();
            var payload = new GoogleDocsApiBatchUpdatePayload();
            AppendDocumentStyleRequests(payload, batch);
            foreach (var request in batch.Requests.Reverse()) {
                AppendRequest(payload, batch.Report, request, imageUris, null, allowFootnotes: true, prepared.Footnotes, allowNamedRanges: true);
            }

            prepared.Payload = payload;
            return prepared;
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildSegmentBatchUpdatePayload(
            GoogleDocsSegment segment,
            TranslationReport report,
            string segmentId,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris) {
            if (segment == null) throw new ArgumentNullException(nameof(segment));
            if (report == null) throw new ArgumentNullException(nameof(report));
            if (string.IsNullOrWhiteSpace(segmentId)) throw new ArgumentException("Segment id is required.", nameof(segmentId));

            var payload = new GoogleDocsApiBatchUpdatePayload();
            foreach (var request in segment.Requests.Reverse()) {
                AppendRequest(payload, report, request, imageUris, segmentId, allowFootnotes: false, null, allowNamedRanges: true);
            }

            return payload;
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildFootnoteBatchUpdatePayload(
            GoogleDocsFootnote footnote,
            TranslationReport report,
            string footnoteId,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris) {
            if (footnote == null) throw new ArgumentNullException(nameof(footnote));
            if (report == null) throw new ArgumentNullException(nameof(report));
            if (string.IsNullOrWhiteSpace(footnoteId)) throw new ArgumentException("Footnote id is required.", nameof(footnoteId));

            var payload = new GoogleDocsApiBatchUpdatePayload();
            foreach (var paragraph in footnote.Paragraphs.Reverse()) {
                AppendParagraphRequests(payload, report, paragraph, imageUris, footnoteId, 1, false, allowFootnotes: false, null, allowNamedRanges: true, sectionStyle: null);
            }

            return payload;
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildTableContentBatchUpdatePayload(
            GoogleDocsBatch batch,
            GoogleDocsApiDocumentResponse documentState,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris = null) {
            return BuildPreparedTableContentBatchUpdate(batch, documentState, imageUris).Payload;
        }

        internal static PreparedTableContentBatchUpdate BuildPreparedTableContentBatchUpdate(
            GoogleDocsBatch batch,
            GoogleDocsApiDocumentResponse documentState,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris = null) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));
            if (documentState == null) throw new ArgumentNullException(nameof(documentState));

            var prepared = new PreparedTableContentBatchUpdate();
            var payload = prepared.Payload;
            var tableRequests = batch.Requests.OfType<GoogleDocsInsertTableRequest>().ToList();
            var documentTables = EnumerateTables(documentState).ToList();
            if (tableRequests.Count != documentTables.Count) {
                batch.Report.Add(
                    TranslationSeverity.Warning,
                    "Tables",
                    $"Expected {tableRequests.Count} Google Docs tables after insertion, but the live document returned {documentTables.Count}. Table cell replay may be incomplete.");
            }

            var cellWrites = new List<(int Index, List<GoogleDocsApiRequestPayload> Requests)>();
            for (int tableIndex = 0; tableIndex < Math.Min(tableRequests.Count, documentTables.Count); tableIndex++) {
                var sourceTable = tableRequests[tableIndex].Table;
                var liveTable = documentTables[tableIndex];

                for (int rowIndex = 0; rowIndex < Math.Min(sourceTable.Rows.Count, liveTable.Rows.Count); rowIndex++) {
                    var sourceRow = sourceTable.Rows[rowIndex];
                    var liveRow = liveTable.Rows[rowIndex];

                    for (int cellIndex = 0; cellIndex < Math.Min(sourceRow.Cells.Count, liveRow.Cells.Count); cellIndex++) {
                        var sourceCell = sourceRow.Cells[cellIndex];
                        var liveCell = liveRow.Cells[cellIndex];
                        if (!TryGetFirstWritableParagraphIndex(liveCell, out var insertionIndex)) {
                            batch.Report.Add(
                                TranslationSeverity.Warning,
                                "Tables",
                                $"Table cell [{rowIndex},{cellIndex}] could not be resolved to a writable Google Docs paragraph index.");
                            continue;
                        }

                        var requests = BuildCellContentRequests(
                            sourceCell,
                            batch.Report,
                            insertionIndex,
                            null,
                            imageUris,
                            allowFootnotes: true,
                            prepared.Footnotes,
                            allowNamedRanges: true);
                        if (!HasMeaningfulCellContentRequests(requests)) {
                            continue;
                        }

                        cellWrites.Add((insertionIndex, requests));
                    }
                }
            }

            foreach (var write in cellWrites.OrderByDescending(item => item.Index)) {
                payload.Requests.AddRange(write.Requests);
            }

            return prepared;
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildSegmentTableContentBatchUpdatePayload(
            GoogleDocsSegment segment,
            GoogleDocsApiDocumentResponse documentState,
            TranslationReport report,
            string segmentId,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris = null) {
            if (segment == null) throw new ArgumentNullException(nameof(segment));
            if (documentState == null) throw new ArgumentNullException(nameof(documentState));
            if (report == null) throw new ArgumentNullException(nameof(report));
            if (string.IsNullOrWhiteSpace(segmentId)) throw new ArgumentException("Segment id is required.", nameof(segmentId));

            var payload = new GoogleDocsApiBatchUpdatePayload();
            var tableRequests = segment.Requests.OfType<GoogleDocsInsertTableRequest>().ToList();
            var documentTables = EnumerateSegmentTables(documentState, segment.Kind, segmentId).ToList();
            if (tableRequests.Count != documentTables.Count) {
                report.Add(
                    TranslationSeverity.Warning,
                    "HeadersAndFooters",
                    $"Expected {tableRequests.Count} Google Docs table(s) inside {segment.Kind} segment '{segmentId}', but the live document returned {documentTables.Count}. Table cell replay may be incomplete.");
            }

            var cellWrites = new List<(int Index, List<GoogleDocsApiRequestPayload> Requests)>();
            for (int tableIndex = 0; tableIndex < Math.Min(tableRequests.Count, documentTables.Count); tableIndex++) {
                var sourceTable = tableRequests[tableIndex].Table;
                var liveTable = documentTables[tableIndex];

                for (int rowIndex = 0; rowIndex < Math.Min(sourceTable.Rows.Count, liveTable.Rows.Count); rowIndex++) {
                    var sourceRow = sourceTable.Rows[rowIndex];
                    var liveRow = liveTable.Rows[rowIndex];

                    for (int cellIndex = 0; cellIndex < Math.Min(sourceRow.Cells.Count, liveRow.Cells.Count); cellIndex++) {
                        var sourceCell = sourceRow.Cells[cellIndex];
                        var liveCell = liveRow.Cells[cellIndex];
                        if (!TryGetFirstWritableParagraphIndex(liveCell, out var insertionIndex)) {
                            report.Add(
                                TranslationSeverity.Warning,
                                "HeadersAndFooters",
                                $"{segment.Kind} table cell [{rowIndex},{cellIndex}] could not be resolved to a writable Google Docs paragraph index.");
                            continue;
                        }

                        var requests = BuildCellContentRequests(
                            sourceCell,
                            report,
                            insertionIndex,
                            segmentId,
                            imageUris,
                            allowFootnotes: false,
                            null,
                            allowNamedRanges: true);
                        if (!HasMeaningfulCellContentRequests(requests)) {
                            continue;
                        }

                        cellWrites.Add((insertionIndex, requests));
                    }
                }
            }

            foreach (var write in cellWrites.OrderByDescending(item => item.Index)) {
                payload.Requests.AddRange(write.Requests);
            }

            return payload;
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildTableMergeBatchUpdatePayload(
            GoogleDocsBatch batch,
            GoogleDocsApiDocumentResponse documentState) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));
            if (documentState == null) throw new ArgumentNullException(nameof(documentState));

            var payload = new GoogleDocsApiBatchUpdatePayload();
            var tableRequests = batch.Requests.OfType<GoogleDocsInsertTableRequest>().ToList();
            var liveTables = EnumerateLiveTables(documentState.Body?.Content).ToList();
            if (tableRequests.Count != liveTables.Count) {
                batch.Report.Add(
                    TranslationSeverity.Warning,
                    "TableMerges",
                    $"Expected {tableRequests.Count} Google Docs tables while preparing merge requests, but the live document returned {liveTables.Count}. Merge replay may be incomplete.");
            }

            for (int tableIndex = 0; tableIndex < Math.Min(tableRequests.Count, liveTables.Count); tableIndex++) {
                AppendMergeRequests(
                    payload,
                    tableRequests[tableIndex].Table,
                    liveTables[tableIndex].StartIndex,
                    null);
            }

            return payload;
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildTableStyleBatchUpdatePayload(
            GoogleDocsBatch batch,
            GoogleDocsApiDocumentResponse documentState) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));
            if (documentState == null) throw new ArgumentNullException(nameof(documentState));

            var payload = new GoogleDocsApiBatchUpdatePayload();
            var tableRequests = batch.Requests.OfType<GoogleDocsInsertTableRequest>().ToList();
            var liveTables = EnumerateLiveTables(documentState.Body?.Content).ToList();
            if (tableRequests.Count != liveTables.Count) {
                batch.Report.Add(
                    TranslationSeverity.Warning,
                    "Tables",
                    $"Expected {tableRequests.Count} Google Docs tables while preparing table style requests, but the live document returned {liveTables.Count}. Table presentation replay may be incomplete.");
            }

            for (int tableIndex = 0; tableIndex < Math.Min(tableRequests.Count, liveTables.Count); tableIndex++) {
                AppendTableStyleRequests(
                    payload,
                    batch.Report,
                    tableRequests[tableIndex].Table,
                    liveTables[tableIndex].StartIndex,
                    null,
                    allowPinnedHeaderRows: true);
            }

            return payload;
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildSegmentTableMergeBatchUpdatePayload(
            GoogleDocsSegment segment,
            GoogleDocsApiDocumentResponse documentState,
            TranslationReport report,
            string segmentId) {
            if (segment == null) throw new ArgumentNullException(nameof(segment));
            if (documentState == null) throw new ArgumentNullException(nameof(documentState));
            if (report == null) throw new ArgumentNullException(nameof(report));
            if (string.IsNullOrWhiteSpace(segmentId)) throw new ArgumentException("Segment id is required.", nameof(segmentId));

            var payload = new GoogleDocsApiBatchUpdatePayload();
            var tableRequests = segment.Requests.OfType<GoogleDocsInsertTableRequest>().ToList();
            var liveTables = EnumerateLiveTables(ResolveSegmentContent(documentState, segment.Kind, segmentId)).ToList();
            if (tableRequests.Count != liveTables.Count) {
                report.Add(
                    TranslationSeverity.Warning,
                    "TableMerges",
                    $"Expected {tableRequests.Count} Google Docs tables while preparing merge requests for {segment.Kind} segment '{segmentId}', but the live document returned {liveTables.Count}. Merge replay may be incomplete.");
            }

            for (int tableIndex = 0; tableIndex < Math.Min(tableRequests.Count, liveTables.Count); tableIndex++) {
                AppendMergeRequests(
                    payload,
                    tableRequests[tableIndex].Table,
                    liveTables[tableIndex].StartIndex,
                    segmentId);
            }

            return payload;
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildSegmentTableStyleBatchUpdatePayload(
            GoogleDocsSegment segment,
            GoogleDocsApiDocumentResponse documentState,
            TranslationReport report,
            string segmentId) {
            if (segment == null) throw new ArgumentNullException(nameof(segment));
            if (documentState == null) throw new ArgumentNullException(nameof(documentState));
            if (report == null) throw new ArgumentNullException(nameof(report));
            if (string.IsNullOrWhiteSpace(segmentId)) throw new ArgumentException("Segment id is required.", nameof(segmentId));

            var payload = new GoogleDocsApiBatchUpdatePayload();
            var tableRequests = segment.Requests.OfType<GoogleDocsInsertTableRequest>().ToList();
            var liveTables = EnumerateLiveTables(ResolveSegmentContent(documentState, segment.Kind, segmentId)).ToList();
            if (tableRequests.Count != liveTables.Count) {
                report.Add(
                    TranslationSeverity.Warning,
                    "HeadersAndFooters",
                    $"Expected {tableRequests.Count} Google Docs tables while preparing table style requests for {segment.Kind} segment '{segmentId}', but the live document returned {liveTables.Count}. Table presentation replay may be incomplete.");
            }

            for (int tableIndex = 0; tableIndex < Math.Min(tableRequests.Count, liveTables.Count); tableIndex++) {
                AppendTableStyleRequests(
                    payload,
                    report,
                    tableRequests[tableIndex].Table,
                    liveTables[tableIndex].StartIndex,
                    segmentId,
                    allowPinnedHeaderRows: false);
            }

            return payload;
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildResetDocumentPayload(GoogleDocsApiDocumentResponse documentState) {
            if (documentState == null) throw new ArgumentNullException(nameof(documentState));

            var payload = new GoogleDocsApiBatchUpdatePayload();
            var bodyEndIndex = GetBodyEndIndex(documentState);
            if (bodyEndIndex <= 2) {
                return payload;
            }

            payload.Requests.Add(new GoogleDocsApiRequestPayload {
                DeleteContentRange = new GoogleDocsApiDeleteContentRangeRequestPayload {
                    Range = new GoogleDocsApiRangePayload {
                        StartIndex = 1,
                        EndIndex = bodyEndIndex - 1,
                    }
                }
            });

            return payload;
        }
    }
}

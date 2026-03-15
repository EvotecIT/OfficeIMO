using OfficeIMO.GoogleWorkspace;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.GoogleDocs {
    internal static class GoogleDocsApiPayloadBuilder {
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
                        var cellText = BuildCellText(sourceCell);
                        if (string.IsNullOrWhiteSpace(cellText)) {
                            continue;
                        }

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
                        if (requests.Count == 0) {
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
                        var cellText = BuildCellText(sourceCell);
                        if (string.IsNullOrWhiteSpace(cellText)) {
                            continue;
                        }

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
                        if (requests.Count == 0) {
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

        private static void AppendParagraphRequests(
            GoogleDocsApiBatchUpdatePayload payload,
            TranslationReport report,
            GoogleDocsParagraph paragraph,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris,
            string? segmentId) {
            AppendParagraphRequests(payload, report, paragraph, imageUris, segmentId, 1, true, segmentId == null, null, allowNamedRanges: segmentId == null, sectionStyle: null);
        }

        private static void AppendParagraphRequests(
            GoogleDocsApiBatchUpdatePayload payload,
            TranslationReport report,
            GoogleDocsParagraph paragraph,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris,
            string? segmentId,
            int insertionIndex,
            bool allowStructuralBreaks,
            bool allowFootnotes,
            List<GoogleDocsFootnote>? preparedFootnotes,
            bool allowNamedRanges,
            GoogleDocsSectionStyle? sectionStyle) {
            var materialized = MaterializeParagraph(paragraph, report, imageUris);
            payload.Requests.Add(new GoogleDocsApiRequestPayload {
                InsertText = new GoogleDocsApiInsertTextRequestPayload {
                    Location = new GoogleDocsApiLocationPayload {
                        Index = insertionIndex,
                        SegmentId = segmentId,
                    },
                    Text = materialized.InsertedText,
                }
            });

            if (!string.IsNullOrWhiteSpace(paragraph.BookmarkName) && allowNamedRanges && materialized.InsertedText.Length > 0) {
                int namedRangeEndIndex = insertionIndex + Math.Max(1, materialized.InsertedText.Length - 1);
                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    CreateNamedRange = new GoogleDocsApiCreateNamedRangeRequestPayload {
                        Name = paragraph.BookmarkName!,
                        Range = new GoogleDocsApiRangePayload {
                            StartIndex = insertionIndex,
                            EndIndex = namedRangeEndIndex,
                            SegmentId = segmentId,
                        }
                    }
                });
            }

            var paragraphFields = new List<string>();
            var paragraphStyle = new GoogleDocsApiParagraphStylePayload();
            if (TryMapNamedStyle(paragraph, out var namedStyleType)) {
                paragraphStyle.NamedStyleType = namedStyleType;
                paragraphFields.Add("namedStyleType");
            }

            if (TryMapAlignment(paragraph.Alignment, out var alignment)) {
                paragraphStyle.Alignment = alignment;
                paragraphFields.Add("alignment");
            }

            if (paragraph.IsRightToLeft) {
                paragraphStyle.Direction = "RIGHT_TO_LEFT";
                paragraphFields.Add("direction");
            }

            if (paragraph.KeepWithNext) {
                paragraphStyle.KeepWithNext = true;
                paragraphFields.Add("keepWithNext");
            }

            if (paragraph.KeepLinesTogether) {
                paragraphStyle.KeepLinesTogether = true;
                paragraphFields.Add("keepLinesTogether");
            }

            if (paragraph.AvoidWidowAndOrphan) {
                paragraphStyle.AvoidWidowAndOrphan = true;
                paragraphFields.Add("avoidWidowAndOrphan");
            }

            if (BuildParagraphTabStops(paragraph.TabStops) is { Count: > 0 } tabStops) {
                paragraphStyle.TabStops = tabStops;
                paragraphFields.Add("tabStops");
            }

            if (TryBuildParagraphDimension(paragraph.IndentStartPoints, out var indentStart)) {
                paragraphStyle.IndentStart = indentStart;
                paragraphFields.Add("indentStart");
            }

            if (TryBuildParagraphDimension(paragraph.IndentEndPoints, out var indentEnd)) {
                paragraphStyle.IndentEnd = indentEnd;
                paragraphFields.Add("indentEnd");
            }

            if (TryBuildParagraphDimension(paragraph.IndentFirstLinePoints, out var indentFirstLine)) {
                paragraphStyle.IndentFirstLine = indentFirstLine;
                paragraphFields.Add("indentFirstLine");
            }

            if (TryBuildParagraphDimension(paragraph.SpaceAbovePoints, out var spaceAbove)) {
                paragraphStyle.SpaceAbove = spaceAbove;
                paragraphFields.Add("spaceAbove");
            }

            if (TryBuildParagraphDimension(paragraph.SpaceBelowPoints, out var spaceBelow)) {
                paragraphStyle.SpaceBelow = spaceBelow;
                paragraphFields.Add("spaceBelow");
            }

            if (paragraph.LineSpacingPercent.HasValue) {
                paragraphStyle.LineSpacing = paragraph.LineSpacingPercent.Value;
                paragraphFields.Add("lineSpacing");
            }

            if (!string.IsNullOrWhiteSpace(paragraph.ShadingFillColorHex)) {
                var shadingColor = BuildOptionalColor(paragraph.ShadingFillColorHex);
                if (shadingColor != null) {
                    paragraphStyle.Shading = new GoogleDocsApiParagraphShadingPayload {
                        BackgroundColor = shadingColor,
                    };
                    paragraphFields.Add("shading");
                }
            }

            if (BuildParagraphBorder(paragraph.LeftBorder) is { } leftBorder) {
                paragraphStyle.BorderLeft = leftBorder;
                paragraphFields.Add("borderLeft");
            }

            if (BuildParagraphBorder(paragraph.RightBorder) is { } rightBorder) {
                paragraphStyle.BorderRight = rightBorder;
                paragraphFields.Add("borderRight");
            }

            if (BuildParagraphBorder(paragraph.TopBorder) is { } topBorder) {
                paragraphStyle.BorderTop = topBorder;
                paragraphFields.Add("borderTop");
            }

            if (BuildParagraphBorder(paragraph.BottomBorder) is { } bottomBorder) {
                paragraphStyle.BorderBottom = bottomBorder;
                paragraphFields.Add("borderBottom");
            }

            if (paragraphFields.Count > 0) {
                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    UpdateParagraphStyle = new GoogleDocsApiUpdateParagraphStyleRequestPayload {
                        Range = new GoogleDocsApiRangePayload {
                            StartIndex = insertionIndex,
                            EndIndex = insertionIndex + materialized.InsertedText.Length,
                            SegmentId = segmentId,
                        },
                        ParagraphStyle = paragraphStyle,
                        Fields = string.Join(",", paragraphFields),
                    }
                });
            }

            AppendSectionStyleRequest(
                payload,
                insertionIndex,
                insertionIndex + materialized.InsertedText.Length,
                segmentId,
                sectionStyle);

            foreach (var run in materialized.Runs) {
                var textStyle = BuildTextStyle(run.Source);
                if (textStyle == null) {
                    continue;
                }

                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    UpdateTextStyle = new GoogleDocsApiUpdateTextStyleRequestPayload {
                        Range = new GoogleDocsApiRangePayload {
                            StartIndex = insertionIndex + run.StartOffset,
                            EndIndex = insertionIndex + run.EndOffset,
                            SegmentId = segmentId,
                        },
                        TextStyle = textStyle,
                        Fields = BuildTextStyleFields(textStyle),
                    }
                });
            }

            if (paragraph.IsListItem && materialized.InsertedText.Length > 1) {
                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    CreateParagraphBullets = new GoogleDocsApiCreateParagraphBulletsRequestPayload {
                        Range = new GoogleDocsApiRangePayload {
                            StartIndex = insertionIndex,
                            EndIndex = insertionIndex + materialized.InsertedText.Length,
                            SegmentId = segmentId,
                        },
                        BulletPreset = ResolveListPreset(paragraph, report),
                    }
                });
            }

            if (materialized.Footnotes.Count > 0 && allowFootnotes) {
                foreach (var footnote in materialized.Footnotes.OrderByDescending(item => item.InsertOffset)) {
                    payload.Requests.Add(new GoogleDocsApiRequestPayload {
                        CreateFootnote = new GoogleDocsApiCreateFootnoteRequestPayload {
                            Location = new GoogleDocsApiLocationPayload {
                                Index = insertionIndex + footnote.InsertOffset,
                                SegmentId = segmentId,
                            }
                        }
                    });
                    preparedFootnotes?.Add(footnote.Source);
                }
            }

            if (materialized.Footnotes.Count > 0 && !allowFootnotes) {
                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Warning,
                    segmentId == null ? "Tables" : "Footnotes",
                    segmentId == null
                        ? "Footnotes inside table cell replay are not executed in the current Google Docs slice, because native footnote creation is currently limited to top-level body paragraphs."
                        : "Nested footnotes inside Google Docs segments are not executed in the current slice.");
            }

            foreach (var image in materialized.Images.OrderByDescending(item => item.InsertOffset)) {
                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    InsertInlineImage = new GoogleDocsApiInsertInlineImageRequestPayload {
                        Uri = image.Uri,
                        ObjectSize = BuildImageSize(image.Source),
                        Location = new GoogleDocsApiLocationPayload {
                            Index = insertionIndex + image.InsertOffset,
                            SegmentId = segmentId,
                        }
                    }
                });
            }

            if (paragraph.PageBreakBefore && segmentId == null && allowStructuralBreaks) {
                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    InsertPageBreak = new GoogleDocsApiInsertPageBreakRequestPayload {
                        Location = new GoogleDocsApiLocationPayload {
                            Index = insertionIndex,
                        }
                    }
                });

                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Info,
                        "PageBreaks",
                        "Word paragraphs marked with PageBreakBefore are now emitted as native Google Docs insertPageBreak requests in the body flow.");
            }

            if (paragraph.PageBreakBefore && segmentId != null) {
                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Warning,
                    "HeadersAndFooters",
                    "PageBreakBefore in a Word header/footer paragraph is not executed in the current Google Docs slice because insertPageBreak is only valid in the document body.");
            }

            if (paragraph.PageBreakBefore && !allowStructuralBreaks) {
                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Warning,
                    segmentId == null ? "Tables" : "HeadersAndFooters",
                    "PageBreakBefore inside a table cell is ignored in the current Google Docs slice because insertPageBreak is only emitted for top-level document flow.");
            }

            if (paragraph.StartsNewSectionBefore && segmentId == null && allowStructuralBreaks) {
                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    InsertSectionBreak = new GoogleDocsApiInsertSectionBreakRequestPayload {
                        SectionType = ResolveSectionBreakType(paragraph.SectionBreakType, report),
                        Location = new GoogleDocsApiLocationPayload {
                            Index = insertionIndex,
                        }
                    }
                });
            }

            if (paragraph.StartsNewSectionBefore && segmentId != null) {
                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Warning,
                    "HeadersAndFooters",
                    "Section-break markers are ignored inside header/footer segment content because insertSectionBreak is only valid in the document body.");
            }

            if (paragraph.StartsNewSectionBefore && !allowStructuralBreaks) {
                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Warning,
                    segmentId == null ? "Tables" : "HeadersAndFooters",
                    "Section-break markers inside a table cell are ignored in the current Google Docs slice because insertSectionBreak is only emitted for top-level document flow.");
            }
        }

        private static void AppendRequest(
            GoogleDocsApiBatchUpdatePayload payload,
            TranslationReport report,
            GoogleDocsRequest request,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris,
            string? segmentId,
            bool allowFootnotes,
            List<GoogleDocsFootnote>? preparedFootnotes,
            bool allowNamedRanges) {
            switch (request) {
                case GoogleDocsInsertParagraphRequest paragraphRequest:
                    AppendParagraphRequests(payload, report, paragraphRequest.Paragraph, imageUris, segmentId, 1, true, allowFootnotes, preparedFootnotes, allowNamedRanges, paragraphRequest.SectionStyle);
                    break;
                case GoogleDocsInsertTableRequest tableRequest:
                    payload.Requests.Add(new GoogleDocsApiRequestPayload {
                        InsertTable = new GoogleDocsApiInsertTableRequestPayload {
                            Rows = Math.Max(1, tableRequest.Table.RowCount),
                            Columns = Math.Max(1, tableRequest.Table.ColumnCount),
                            Location = new GoogleDocsApiLocationPayload {
                                Index = 1,
                                SegmentId = segmentId,
                            }
                        }
                    });

                    if (tableRequest.StartsNewSectionBefore && segmentId == null) {
                        payload.Requests.Add(new GoogleDocsApiRequestPayload {
                            InsertSectionBreak = new GoogleDocsApiInsertSectionBreakRequestPayload {
                        SectionType = ResolveSectionBreakType(tableRequest.SectionBreakType, report),
                        Location = new GoogleDocsApiLocationPayload {
                            Index = 1,
                        }
                            }
                        });
                    }

                    if (tableRequest.StartsNewSectionBefore && segmentId != null) {
                        AddReportNoticeOnce(
                            report,
                            TranslationSeverity.Warning,
                            "HeadersAndFooters",
                            "Section-break markers are ignored inside header/footer segment content because insertSectionBreak is only valid in the document body.");
                    }

                    AppendSectionStyleRequest(payload, 1, 2, segmentId, tableRequest.SectionStyle);
                    break;
            }
        }

        private static void AppendMergeRequests(
            GoogleDocsApiBatchUpdatePayload payload,
            GoogleDocsTable table,
            int tableStartIndex,
            string? segmentId) {
            foreach (var row in table.Rows) {
                foreach (var cell in row.Cells) {
                    if (cell.ColumnSpan <= 1 && cell.RowSpan <= 1) {
                        continue;
                    }

                    payload.Requests.Add(new GoogleDocsApiRequestPayload {
                        MergeTableCells = new GoogleDocsApiMergeTableCellsRequestPayload {
                            TableRange = new GoogleDocsApiTableRangePayload {
                                RowSpan = Math.Max(1, cell.RowSpan),
                                ColumnSpan = Math.Max(1, cell.ColumnSpan),
                                TableCellLocation = new GoogleDocsApiTableCellLocationPayload {
                                    RowIndex = row.RowIndex,
                                    ColumnIndex = cell.ColumnIndex,
                                    TableStartLocation = new GoogleDocsApiLocationPayload {
                                        Index = tableStartIndex,
                                        SegmentId = segmentId,
                                    }
                                }
                            }
                        }
                    });
                }
            }
        }

        private static void AppendTableStyleRequests(
            GoogleDocsApiBatchUpdatePayload payload,
            TranslationReport report,
            GoogleDocsTable table,
            int tableStartIndex,
            string? segmentId,
            bool allowPinnedHeaderRows) {
            if (allowPinnedHeaderRows && table.RepeatHeaderRow) {
                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    PinTableHeaderRows = new GoogleDocsApiPinTableHeaderRowsRequestPayload {
                        PinnedHeaderRowsCount = 1,
                        TableStartLocation = new GoogleDocsApiLocationPayload {
                            Index = tableStartIndex,
                            SegmentId = segmentId,
                        }
                    }
                });

                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Info,
                    segmentId == null ? "Tables" : "HeadersAndFooters",
                    "Word table header rows are now emitted as native Google Docs pinTableHeaderRows requests when the live table structure can be resolved.");
            }

            for (int columnIndex = 0; columnIndex < table.ColumnWidthPoints.Count; columnIndex++) {
                var widthPoints = table.ColumnWidthPoints[columnIndex];
                if (widthPoints <= 0) {
                    continue;
                }

                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    UpdateTableColumnProperties = new GoogleDocsApiUpdateTableColumnPropertiesRequestPayload {
                        TableStartLocation = new GoogleDocsApiLocationPayload {
                            Index = tableStartIndex,
                            SegmentId = segmentId,
                        },
                        ColumnIndices = new List<int> { columnIndex },
                        TableColumnProperties = new GoogleDocsApiTableColumnPropertiesPayload {
                            WidthType = "FIXED_WIDTH",
                            Width = new GoogleDocsApiDimensionPayload {
                                Magnitude = Math.Round(widthPoints, 2, MidpointRounding.AwayFromZero),
                                Unit = "PT",
                            },
                        },
                        Fields = "width,widthType",
                    }
                });
            }

            if (table.ColumnWidthPoints.Count > 0) {
                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Info,
                    segmentId == null ? "Tables" : "HeadersAndFooters",
                    "Word table column widths are now emitted as native Google Docs updateTableColumnProperties requests when the live table structure can be resolved.");
            }

            foreach (var row in table.Rows) {
                foreach (var cell in row.Cells) {
                    if ((cell.HasHorizontalMerge && cell.ColumnSpan <= 1) || (cell.HasVerticalMerge && cell.RowSpan <= 1)) {
                        continue;
                    }

                    var tableCellStyle = new GoogleDocsApiTableCellStylePayload {
                        BackgroundColor = BuildOptionalColor(cell.ShadingFillColorHex),
                        BorderLeft = BuildTableCellBorder(cell.LeftBorder),
                        BorderRight = BuildTableCellBorder(cell.RightBorder),
                        BorderTop = BuildTableCellBorder(cell.TopBorder),
                        BorderBottom = BuildTableCellBorder(cell.BottomBorder),
                    };

                    var fields = BuildTableCellStyleFields(tableCellStyle);
                    if (fields.Count == 0) {
                        continue;
                    }

                    payload.Requests.Add(new GoogleDocsApiRequestPayload {
                        UpdateTableCellStyle = new GoogleDocsApiUpdateTableCellStyleRequestPayload {
                            Fields = string.Join(",", fields),
                            TableCellStyle = tableCellStyle,
                            TableRange = new GoogleDocsApiTableRangePayload {
                                RowSpan = Math.Max(1, cell.RowSpan),
                                ColumnSpan = Math.Max(1, cell.ColumnSpan),
                                TableCellLocation = new GoogleDocsApiTableCellLocationPayload {
                                    RowIndex = row.RowIndex,
                                    ColumnIndex = cell.ColumnIndex,
                                    TableStartLocation = new GoogleDocsApiLocationPayload {
                                        Index = tableStartIndex,
                                        SegmentId = segmentId,
                                    }
                                }
                            }
                        }
                    });

                    AddReportNoticeOnce(
                        report,
                        TranslationSeverity.Info,
                        segmentId == null ? "Tables" : "HeadersAndFooters",
                        "Word table cell shading and supported border edges are now emitted as native Google Docs updateTableCellStyle requests when the live table structure can be resolved.");
                }
            }
        }

        private static List<string> BuildTableCellStyleFields(GoogleDocsApiTableCellStylePayload style) {
            var fields = new List<string>();
            if (style.BackgroundColor != null) fields.Add("backgroundColor");
            if (style.BorderLeft != null) fields.Add("borderLeft");
            if (style.BorderRight != null) fields.Add("borderRight");
            if (style.BorderTop != null) fields.Add("borderTop");
            if (style.BorderBottom != null) fields.Add("borderBottom");
            return fields;
        }

        private static List<GoogleDocsApiRequestPayload> BuildCellContentRequests(
            GoogleDocsTableCell cell,
            TranslationReport report,
            int insertionIndex,
            string? segmentId,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris,
            bool allowFootnotes,
            List<GoogleDocsFootnote>? preparedFootnotes,
            bool allowNamedRanges) {
            var payload = new GoogleDocsApiBatchUpdatePayload();
            foreach (var paragraph in cell.Paragraphs.AsEnumerable().Reverse()) {
                AppendParagraphRequests(payload, report, paragraph, imageUris, segmentId, insertionIndex, false, allowFootnotes, preparedFootnotes, allowNamedRanges, sectionStyle: null);
            }

            return payload.Requests;
        }

        private static MaterializedParagraph MaterializeParagraph(
            GoogleDocsParagraph paragraph,
            TranslationReport report,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris) {
            var materialized = new MaterializedParagraph();
            materialized.PrefixText = BuildParagraphPrefix(paragraph);
            var currentOffset = materialized.PrefixText.Length;
            foreach (var run in paragraph.Runs) {
                var runText = BuildRunText(run, report, imageUris, materialized, currentOffset);
                if (string.IsNullOrEmpty(runText)) {
                    continue;
                }

                materialized.TextBuilder.Append(runText);
                materialized.Runs.Add(new MaterializedRun {
                    StartOffset = currentOffset,
                    EndOffset = currentOffset + runText.Length,
                    Source = run,
                });
                currentOffset += runText.Length;
            }

            if (materialized.TextBuilder.Length == 0 && !string.IsNullOrWhiteSpace(paragraph.Text)) {
                var paragraphText = SanitizeText(paragraph.Text);
                materialized.TextBuilder.Append(paragraphText);
                materialized.Runs.Add(new MaterializedRun {
                    StartOffset = currentOffset,
                    EndOffset = currentOffset + paragraphText.Length,
                    Source = new GoogleDocsParagraphRun {
                        Text = paragraphText,
                    }
                });
            }

            materialized.TextBuilder.Append('\n');
            return materialized;
        }

        private static GoogleDocsApiTextStylePayload? BuildTextStyle(GoogleDocsParagraphRun source) {
            var style = new GoogleDocsApiTextStylePayload();
            var hasStyle = false;

            if (source.Bold) {
                style.Bold = true;
                hasStyle = true;
            }

            if (source.Italic) {
                style.Italic = true;
                hasStyle = true;
            }

            if (source.Underline) {
                style.Underline = true;
                hasStyle = true;
            }

            if (source.Strike) {
                style.Strikethrough = true;
                hasStyle = true;
            }

            if (source.FontSize.HasValue) {
                style.FontSize = new GoogleDocsApiDimensionPayload {
                    Magnitude = source.FontSize.Value,
                    Unit = "PT",
                };
                hasStyle = true;
            }

            if (!string.IsNullOrWhiteSpace(source.FontFamily)) {
                style.WeightedFontFamily = new GoogleDocsApiWeightedFontFamilyPayload {
                    FontFamily = source.FontFamily!.Trim(),
                };
                hasStyle = true;
            }

            if (!string.IsNullOrWhiteSpace(source.ForegroundColorHex)) {
                style.ForegroundColor = BuildOptionalColor(source.ForegroundColorHex);
                hasStyle = style.ForegroundColor != null;
            }

            if (!string.IsNullOrWhiteSpace(source.HighlightColor)) {
                style.BackgroundColor = BuildHighlightColor(source.HighlightColor);
                hasStyle |= style.BackgroundColor != null;
            }

            if (!string.IsNullOrWhiteSpace(source.VerticalTextAlignment)) {
                style.BaselineOffset = BuildBaselineOffset(source.VerticalTextAlignment);
                hasStyle |= !string.IsNullOrWhiteSpace(style.BaselineOffset);
            }

            if (string.Equals(source.CapsStyle, "SmallCaps", StringComparison.OrdinalIgnoreCase)) {
                style.SmallCaps = true;
                hasStyle = true;
            }

            if (!string.IsNullOrWhiteSpace(source.Link?.Uri)) {
                var linkUri = source.Link!.Uri!;
                style.Link = new GoogleDocsApiLinkPayload {
                    Url = linkUri,
                };
                hasStyle = true;
            }

            return hasStyle ? style : null;
        }

        private static string BuildTextStyleFields(GoogleDocsApiTextStylePayload style) {
            var fields = new List<string>();
            if (style.Bold.HasValue) fields.Add("bold");
            if (style.Italic.HasValue) fields.Add("italic");
            if (style.Underline.HasValue) fields.Add("underline");
            if (style.Strikethrough.HasValue) fields.Add("strikethrough");
            if (style.FontSize != null) fields.Add("fontSize");
            if (style.WeightedFontFamily != null) fields.Add("weightedFontFamily");
            if (style.ForegroundColor != null) fields.Add("foregroundColor");
            if (style.BackgroundColor != null) fields.Add("backgroundColor");
            if (!string.IsNullOrWhiteSpace(style.BaselineOffset)) fields.Add("baselineOffset");
            if (style.SmallCaps.HasValue) fields.Add("smallCaps");
            if (style.Link != null) fields.Add("link");
            return string.Join(",", fields);
        }

        private static string BuildRunText(
            GoogleDocsParagraphRun run,
            TranslationReport report,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris,
            MaterializedParagraph materialized,
            int currentOffset) {
            var text = SanitizeText(run.Text);
            if (run.Footnote != null) {
                materialized.Footnotes.Add(new MaterializedFootnote {
                    InsertOffset = currentOffset + text.Length,
                    Source = run.Footnote,
                });
            }

            if (run.InlineImage == null) {
                return text;
            }

            if (imageUris != null && imageUris.TryGetValue(run.InlineImage, out var uploadedUri) && !string.IsNullOrWhiteSpace(uploadedUri)) {
                materialized.Images.Add(new MaterializedImage {
                    InsertOffset = currentOffset + text.Length,
                    Uri = uploadedUri,
                    Source = run.InlineImage,
                });

                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Info,
                    "InlineImages",
                    "Inline Word images are exported through temporary Drive-backed public URLs so Google Docs can materialize native inline images.");

                return text;
            }

            AddReportNoticeOnce(
                report,
                TranslationSeverity.Info,
                "InlineImages",
                "Inline Word images without an uploaded Drive-backed URI fall back to readable placeholders in Google Docs export.");

            var label = run.InlineImage.Description ?? run.InlineImage.Title;
            if (string.IsNullOrWhiteSpace(label) && !string.IsNullOrWhiteSpace(run.InlineImage.FilePath)) {
                label = Path.GetFileName(run.InlineImage.FilePath);
            }

            var placeholder = string.IsNullOrWhiteSpace(label) ? "[Image]" : "[Image: " + label!.Trim() + "]";
            return string.IsNullOrWhiteSpace(text) ? placeholder : text + placeholder;
        }

        private static string BuildParagraphPrefix(GoogleDocsParagraph paragraph) {
            if (!paragraph.IsListItem) {
                return string.Empty;
            }

            var level = paragraph.ListLevel.GetValueOrDefault();
            if (level <= 0) {
                return string.Empty;
            }

            return new string('\t', level);
        }

        private static string ResolveListPreset(
            GoogleDocsParagraph paragraph,
            TranslationReport report) {
            if (paragraph.IsOrderedList == true) {
                return "NUMBERED_DECIMAL_NESTED";
            }

            if (paragraph.IsOrderedList == false) {
                return "BULLET_DISC_CIRCLE_SQUARE";
            }

            AddReportNoticeOnce(
                report,
                TranslationSeverity.Warning,
                "Lists",
                "A Word list item did not expose an ordered-vs-bulleted classification, so Google Docs export currently falls back to a bullet preset for that paragraph.");

            return "BULLET_DISC_CIRCLE_SQUARE";
        }

        private static bool TryBuildParagraphDimension(double? points, out GoogleDocsApiDimensionPayload? dimension) {
            if (!points.HasValue) {
                dimension = null;
                return false;
            }

            dimension = new GoogleDocsApiDimensionPayload {
                Magnitude = points.Value,
                Unit = "PT",
            };
            return true;
        }

        private static string ResolveSectionBreakType(
            string? sectionBreakType,
            TranslationReport report) {
            if (string.IsNullOrWhiteSpace(sectionBreakType)) {
                return "NEXT_PAGE";
            }

            switch (sectionBreakType!.Trim().ToUpperInvariant()) {
                case "CONTINUOUS":
                    return "CONTINUOUS";
                case "NEXTPAGE":
                    return "NEXT_PAGE";
                case "EVENPAGE":
                case "ODDPAGE":
                    AddReportNoticeOnce(
                        report,
                        TranslationSeverity.Warning,
                        "SectionBreaks",
                        "Word even-page and odd-page section breaks currently fall back to Google Docs NEXT_PAGE section breaks.");
                    return "NEXT_PAGE";
                default:
                    AddReportNoticeOnce(
                        report,
                        TranslationSeverity.Warning,
                        "SectionBreaks",
                        $"Word section break type '{sectionBreakType}' is not mapped directly yet, so Google Docs export currently falls back to NEXT_PAGE.");
                    return "NEXT_PAGE";
            }
        }

        private static GoogleDocsApiSizePayload? BuildImageSize(GoogleDocsInlineImage image) {
            var width = TryConvertImageDimension(image.Width);
            var height = TryConvertImageDimension(image.Height);
            if (width == null && height == null) {
                return null;
            }

            return new GoogleDocsApiSizePayload {
                Width = width,
                Height = height,
            };
        }

        private static GoogleDocsApiDimensionPayload? TryConvertImageDimension(double? value) {
            if (!value.HasValue || value.Value <= 0) {
                return null;
            }

            // OfficeIMO image dimensions are authored in inches, so translate to Google Docs points.
            return new GoogleDocsApiDimensionPayload {
                Magnitude = Math.Round(value.Value * 72d, 2),
                Unit = "PT",
            };
        }

        private static string BuildCellText(GoogleDocsTableCell cell) {
            var paragraphs = cell.Paragraphs
                .Select(paragraph => {
                    var text = string.Concat(paragraph.Runs.Select(run => SanitizeText(run.Text)));
                    if (string.IsNullOrWhiteSpace(text)) {
                        text = SanitizeText(paragraph.Text);
                    }
                    return text;
                })
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();

            return paragraphs.Count == 0 ? string.Empty : string.Join("\n", paragraphs);
        }

        private static int GetBodyEndIndex(GoogleDocsApiDocumentResponse documentState) {
            return documentState.Body?.Content?.Select(element => element.EndIndex ?? 0).DefaultIfEmpty(2).Max() ?? 2;
        }

        private static IEnumerable<GoogleDocsApiTableResponse> EnumerateTables(GoogleDocsApiDocumentResponse documentState) {
            if (documentState.Body?.Content == null) {
                yield break;
            }

            foreach (var element in documentState.Body.Content) {
                if (element.Table != null) {
                    yield return element.Table;
                }
            }
        }

        private static IEnumerable<LiveTableContext> EnumerateLiveTables(List<GoogleDocsApiStructuralElementResponse>? content) {
            if (content == null) {
                yield break;
            }

            foreach (var element in content) {
                if (element.Table != null && element.StartIndex.HasValue) {
                    yield return new LiveTableContext {
                        StartIndex = element.StartIndex.Value,
                        Table = element.Table,
                    };
                }
            }
        }

        private static IEnumerable<GoogleDocsApiTableResponse> EnumerateSegmentTables(
            GoogleDocsApiDocumentResponse documentState,
            string kind,
            string segmentId) {
            var content = ResolveSegmentContent(documentState, kind, segmentId);
            if (content == null) {
                yield break;
            }

            foreach (var element in content) {
                if (element.Table != null) {
                    yield return element.Table;
                }
            }
        }

        private static List<GoogleDocsApiStructuralElementResponse>? ResolveSegmentContent(
            GoogleDocsApiDocumentResponse documentState,
            string kind,
            string segmentId) {
            if (string.Equals(kind, "header", StringComparison.OrdinalIgnoreCase)) {
                return documentState.Headers != null && documentState.Headers.TryGetValue(segmentId, out var header)
                    ? header.Content
                    : null;
            }

            if (string.Equals(kind, "footer", StringComparison.OrdinalIgnoreCase)) {
                return documentState.Footers != null && documentState.Footers.TryGetValue(segmentId, out var footer)
                    ? footer.Content
                    : null;
            }

            return null;
        }

        private static bool TryGetFirstWritableParagraphIndex(GoogleDocsApiTableCellResponse cell, out int index) {
            index = 0;
            var paragraphElement = cell.Content?.FirstOrDefault(element => element.Paragraph != null && element.StartIndex.HasValue);
            if (paragraphElement?.StartIndex is int paragraphStartIndex && paragraphStartIndex > 0) {
                index = paragraphStartIndex;
                return true;
            }

            var structuralStartIndex = cell.Content?.FirstOrDefault(element => element.StartIndex.HasValue)?.StartIndex;
            if (structuralStartIndex.HasValue && structuralStartIndex.Value > 0) {
                index = structuralStartIndex.Value;
                return true;
            }

            return false;
        }

        private static bool TryMapNamedStyle(GoogleDocsParagraph paragraph, out string namedStyleType) {
            namedStyleType = string.Empty;
            var value = paragraph.StyleId ?? paragraph.StyleName;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            var normalizedValue = value!.Replace(" ", string.Empty).ToUpperInvariant();
            switch (normalizedValue) {
                case "TITLE":
                    namedStyleType = "TITLE";
                    return true;
                case "SUBTITLE":
                    namedStyleType = "SUBTITLE";
                    return true;
                case "NORMAL":
                case "NORMALTEXT":
                    namedStyleType = "NORMAL_TEXT";
                    return true;
                case "HEADING1":
                    namedStyleType = "HEADING_1";
                    return true;
                case "HEADING2":
                    namedStyleType = "HEADING_2";
                    return true;
                case "HEADING3":
                    namedStyleType = "HEADING_3";
                    return true;
                case "HEADING4":
                    namedStyleType = "HEADING_4";
                    return true;
                case "HEADING5":
                    namedStyleType = "HEADING_5";
                    return true;
                case "HEADING6":
                    namedStyleType = "HEADING_6";
                    return true;
                default:
                    return false;
            }
        }

        private static bool TryMapAlignment(string? alignment, out string docsAlignment) {
            docsAlignment = string.Empty;
            if (string.IsNullOrWhiteSpace(alignment)) {
                return false;
            }

            var normalizedAlignment = alignment!.Trim().ToUpperInvariant();
            switch (normalizedAlignment) {
                case "CENTER":
                    docsAlignment = "CENTER";
                    return true;
                case "BOTH":
                case "JUSTIFIED":
                    docsAlignment = "JUSTIFIED";
                    return true;
                case "RIGHT":
                case "END":
                    docsAlignment = "END";
                    return true;
                case "LEFT":
                case "START":
                    docsAlignment = "START";
                    return true;
                default:
                    return false;
            }
        }

        private static GoogleDocsApiOptionalColorPayload? BuildOptionalColor(string? colorHex) {
            if (!TryParseRgbColor(colorHex, out var red, out var green, out var blue)) {
                return null;
            }

            return new GoogleDocsApiOptionalColorPayload {
                Color = new GoogleDocsApiColorPayload {
                    RgbColor = new GoogleDocsApiRgbColorPayload {
                        Red = red,
                        Green = green,
                        Blue = blue,
                    }
                }
            };
        }

        private static GoogleDocsApiOptionalColorPayload? BuildHighlightColor(string? highlightColor) {
            if (string.IsNullOrWhiteSpace(highlightColor)) {
                return null;
            }

            string nonNullHighlightColor = highlightColor!;
            var normalizedHighlightColor = nonNullHighlightColor.Trim().ToUpperInvariant();
            string? colorHex = normalizedHighlightColor switch {
                "YELLOW" => "FFFF00",
                "GREEN" => "00FF00",
                "CYAN" => "00FFFF",
                "MAGENTA" => "FF00FF",
                "BLUE" => "0000FF",
                "RED" => "FF0000",
                "DARKBLUE" => "000080",
                "DARKCYAN" => "008080",
                "DARKGREEN" => "008000",
                "DARKMAGENTA" => "800080",
                "DARKRED" => "800000",
                "DARKYELLOW" => "808000",
                "DARKGRAY" => "808080",
                "LIGHTGRAY" => "D3D3D3",
                "BLACK" => "000000",
                "WHITE" => "FFFFFF",
                "NONE" => null,
                _ => null,
            };

            return BuildOptionalColor(colorHex);
        }

        private static string? BuildBaselineOffset(string? verticalTextAlignment) {
            if (string.IsNullOrWhiteSpace(verticalTextAlignment)) {
                return null;
            }

            string nonNullVerticalTextAlignment = verticalTextAlignment!;
            return nonNullVerticalTextAlignment.Trim().ToUpperInvariant() switch {
                "SUPERSCRIPT" => "SUPERSCRIPT",
                "SUBSCRIPT" => "SUBSCRIPT",
                _ => null,
            };
        }

        private static GoogleDocsApiTableCellBorderPayload? BuildTableCellBorder(GoogleDocsTableCellBorder? border) {
            if (border == null) {
                return null;
            }

            bool isExplicitlyNone = string.Equals(border.Style, "Nil", StringComparison.OrdinalIgnoreCase)
                || string.Equals(border.Style, "None", StringComparison.OrdinalIgnoreCase);

            var width = BuildBorderWidth(border.Size, isExplicitlyNone);
            var color = BuildOptionalColor(border.ColorHex);
            var dashStyle = ResolveTableBorderDashStyle(border.Style);
            if (width == null && color == null && dashStyle == null) {
                return null;
            }

            return new GoogleDocsApiTableCellBorderPayload {
                Width = width,
                Color = color,
                DashStyle = dashStyle,
            };
        }

        private static GoogleDocsApiParagraphBorderPayload? BuildParagraphBorder(GoogleDocsParagraphBorder? border) {
            if (border == null) {
                return null;
            }

            bool isExplicitlyNone = string.Equals(border.Style, "Nil", StringComparison.OrdinalIgnoreCase)
                || string.Equals(border.Style, "None", StringComparison.OrdinalIgnoreCase);

            var width = BuildBorderWidth(border.Size, isExplicitlyNone);
            var color = BuildOptionalColor(border.ColorHex);
            var padding = BuildParagraphBorderPadding(border.Space);
            var dashStyle = ResolveTableBorderDashStyle(border.Style);
            if (width == null && color == null && padding == null && dashStyle == null) {
                return null;
            }

            return new GoogleDocsApiParagraphBorderPayload {
                Width = width,
                Color = color,
                Padding = padding,
                DashStyle = dashStyle,
            };
        }

        private static void AppendSectionStyleRequest(
            GoogleDocsApiBatchUpdatePayload payload,
            int startIndex,
            int endIndex,
            string? segmentId,
            GoogleDocsSectionStyle? sectionStyle) {
            if (payload == null || sectionStyle == null || !string.IsNullOrWhiteSpace(segmentId)) {
                return;
            }

            var stylePayload = BuildSectionStyle(sectionStyle);
            if (stylePayload == null) {
                return;
            }

            var fields = BuildSectionStyleFields(stylePayload);
            if (fields.Count == 0) {
                return;
            }

            payload.Requests.Add(new GoogleDocsApiRequestPayload {
                UpdateSectionStyle = new GoogleDocsApiUpdateSectionStyleRequestPayload {
                    Range = new GoogleDocsApiRangePayload {
                        StartIndex = startIndex,
                        EndIndex = Math.Max(startIndex + 1, endIndex),
                    },
                    SectionStyle = stylePayload,
                    Fields = string.Join(",", fields),
                }
            });
        }

        private static GoogleDocsApiSectionStylePayload? BuildSectionStyle(GoogleDocsSectionStyle source) {
            if (source == null) {
                return null;
            }

            var payload = new GoogleDocsApiSectionStylePayload();
            if (TryBuildSize(source.PageWidthPoints, source.PageHeightPoints, out var pageSize)) {
                payload.PageSize = pageSize;
            }

            if (TryBuildParagraphDimension(source.MarginTopPoints, out var marginTop)) {
                payload.MarginTop = marginTop;
            }

            if (TryBuildParagraphDimension(source.MarginBottomPoints, out var marginBottom)) {
                payload.MarginBottom = marginBottom;
            }

            if (TryBuildParagraphDimension(source.MarginLeftPoints, out var marginLeft)) {
                payload.MarginLeft = marginLeft;
            }

            if (TryBuildParagraphDimension(source.MarginRightPoints, out var marginRight)) {
                payload.MarginRight = marginRight;
            }

            if (TryBuildParagraphDimension(source.HeaderMarginPoints, out var marginHeader)) {
                payload.MarginHeader = marginHeader;
            }

            if (TryBuildParagraphDimension(source.FooterMarginPoints, out var marginFooter)) {
                payload.MarginFooter = marginFooter;
            }

            if (BuildSectionColumnProperties(source.ColumnCount, source.ColumnSpacingPoints) is { Count: > 0 } columnProperties) {
                payload.ColumnProperties = columnProperties;
            }

            if (source.ColumnCount.GetValueOrDefault() > 1 || source.HasColumnSeparator) {
                payload.ColumnSeparatorStyle = source.HasColumnSeparator ? "BETWEEN_EACH_COLUMN" : "NONE";
            }

            if (source.UseFirstPageHeaderFooter) {
                payload.UseFirstPageHeaderFooter = true;
            }

            if (source.PageNumberStart.HasValue && source.PageNumberStart.Value > 0) {
                payload.PageNumberStart = source.PageNumberStart.Value;
            }

            if (string.Equals(source.Orientation, "Landscape", StringComparison.OrdinalIgnoreCase)) {
                payload.FlipPageOrientation = true;
            }

            return payload;
        }

        private static List<string> BuildSectionStyleFields(GoogleDocsApiSectionStylePayload style) {
            var fields = new List<string>();
            if (style.PageSize != null) fields.Add("pageSize");
            if (style.MarginTop != null) fields.Add("marginTop");
            if (style.MarginBottom != null) fields.Add("marginBottom");
            if (style.MarginLeft != null) fields.Add("marginLeft");
            if (style.MarginRight != null) fields.Add("marginRight");
            if (style.MarginHeader != null) fields.Add("marginHeader");
            if (style.MarginFooter != null) fields.Add("marginFooter");
            if (style.ColumnProperties != null) fields.Add("columnProperties");
            if (!string.IsNullOrWhiteSpace(style.ColumnSeparatorStyle)) fields.Add("columnSeparatorStyle");
            if (style.UseFirstPageHeaderFooter.HasValue) fields.Add("useFirstPageHeaderFooter");
            if (style.PageNumberStart.HasValue) fields.Add("pageNumberStart");
            if (style.FlipPageOrientation.HasValue) fields.Add("flipPageOrientation");
            return fields;
        }

        private static bool TryBuildSize(double? widthPoints, double? heightPoints, out GoogleDocsApiSizePayload size) {
            size = null!;
            if (!widthPoints.HasValue || !heightPoints.HasValue || widthPoints.Value <= 0 || heightPoints.Value <= 0) {
                return false;
            }

            size = new GoogleDocsApiSizePayload {
                Width = new GoogleDocsApiDimensionPayload {
                    Magnitude = Math.Round(widthPoints.Value, 2, MidpointRounding.AwayFromZero),
                    Unit = "PT",
                },
                Height = new GoogleDocsApiDimensionPayload {
                    Magnitude = Math.Round(heightPoints.Value, 2, MidpointRounding.AwayFromZero),
                    Unit = "PT",
                },
            };
            return true;
        }

        private static void AppendDocumentStyleRequests(
            GoogleDocsApiBatchUpdatePayload payload,
            GoogleDocsBatch batch) {
            if (payload == null || batch == null) {
                return;
            }

            bool useEvenPageHeaderFooter = batch.Snapshot.Sections.Any(section => section.DifferentOddAndEvenPages)
                || batch.Segments.Any(segment => string.Equals(segment.Variant, "even", StringComparison.OrdinalIgnoreCase));
            if (!useEvenPageHeaderFooter) {
                return;
            }

            payload.Requests.Add(new GoogleDocsApiRequestPayload {
                UpdateDocumentStyle = new GoogleDocsApiUpdateDocumentStyleRequestPayload {
                    DocumentStyle = new GoogleDocsApiDocumentStylePayload {
                        UseEvenPageHeaderFooter = true,
                    },
                    Fields = "useEvenPageHeaderFooter",
                }
            });
        }

        private static List<GoogleDocsApiSectionColumnPropertiesPayload>? BuildSectionColumnProperties(int? columnCount, double? columnSpacingPoints) {
            int count = columnCount.GetValueOrDefault();
            if (count <= 1) {
                return null;
            }

            var columns = new List<GoogleDocsApiSectionColumnPropertiesPayload>();
            for (int index = 0; index < count; index++) {
                var column = new GoogleDocsApiSectionColumnPropertiesPayload();
                if (index < count - 1 && TryBuildParagraphDimension(columnSpacingPoints, out var paddingEnd)) {
                    column.PaddingEnd = paddingEnd;
                }

                columns.Add(column);
            }

            return columns;
        }

        private static List<GoogleDocsApiTabStopPayload>? BuildParagraphTabStops(IReadOnlyList<GoogleDocsTabStop> tabStops) {
            if (tabStops == null || tabStops.Count == 0) {
                return null;
            }

            var result = new List<GoogleDocsApiTabStopPayload>();
            foreach (var tabStop in tabStops) {
                if (!TryBuildParagraphTabStop(tabStop, out var payload)) {
                    continue;
                }

                result.Add(payload);
            }

            return result.Count == 0 ? null : result;
        }

        private static bool TryBuildParagraphTabStop(GoogleDocsTabStop tabStop, out GoogleDocsApiTabStopPayload payload) {
            payload = null!;
            if (tabStop == null || tabStop.OffsetPoints < 0) {
                return false;
            }

            payload = new GoogleDocsApiTabStopPayload {
                Alignment = ResolveParagraphTabStopAlignment(tabStop.Alignment),
                Offset = new GoogleDocsApiDimensionPayload {
                    Magnitude = Math.Round(tabStop.OffsetPoints, 2, MidpointRounding.AwayFromZero),
                    Unit = "PT",
                }
            };

            return true;
        }

        private static string? ResolveParagraphTabStopAlignment(string? alignment) {
            if (string.IsNullOrWhiteSpace(alignment)) {
                return null;
            }

            switch (alignment!.Trim().ToUpperInvariant()) {
                case "LEFT":
                case "START":
                case "BAR":
                case "CLEAR":
                case "LIST":
                    return "START";
                case "CENTER":
                    return "CENTER";
                case "RIGHT":
                case "END":
                    return "END";
                case "DECIMAL":
                    return "DECIMAL";
                default:
                    return "START";
            }
        }

        private static GoogleDocsApiDimensionPayload? BuildBorderWidth(uint? size, bool isExplicitlyNone) {
            if (isExplicitlyNone) {
                return new GoogleDocsApiDimensionPayload {
                    Magnitude = 0,
                    Unit = "PT",
                };
            }

            if (!size.HasValue || size.Value == 0) {
                return null;
            }

            return new GoogleDocsApiDimensionPayload {
                Magnitude = Math.Round(size.Value / 8d, 2, MidpointRounding.AwayFromZero),
                Unit = "PT",
            };
        }

        private static GoogleDocsApiDimensionPayload? BuildParagraphBorderPadding(uint? space) {
            if (!space.HasValue) {
                return null;
            }

            return new GoogleDocsApiDimensionPayload {
                Magnitude = space.Value,
                Unit = "PT",
            };
        }

        private static string? ResolveTableBorderDashStyle(string? style) {
            if (string.IsNullOrWhiteSpace(style)) {
                return null;
            }

            switch (style!.Trim().ToUpperInvariant()) {
                case "NONE":
                case "NIL":
                case "SINGLE":
                case "THICK":
                case "DOUBLE":
                case "TRIPLE":
                case "THINTHICKSMALLGAP":
                case "THICKTHINSMALLGAP":
                case "THINTHICKTHINSMALLGAP":
                case "THINTHICKMEDIUMGAP":
                case "THICKTHINMEDIUMGAP":
                case "THINTHICKTHINMEDIUMGAP":
                case "THINTHICKLARGEGAP":
                case "THICKTHINLARGEGAP":
                case "THINTHICKTHINLARGEGAP":
                case "WAVE":
                case "DOUBLEWAVE":
                case "THREED":
                case "THREEDEMBOSS":
                case "THREEDENGRAVE":
                case "OUTSET":
                case "INSET":
                    return "SOLID";
                case "DASHDOT":
                case "DASHDOTSTROKED":
                case "DOTDASH":
                case "DOTDOTDASH":
                    return "DASH_DOT";
                case "DASH":
                case "DASHED":
                case "DASHSMALLGAP":
                case "DASHDOTDOTHEAVY":
                case "DASHDOTHEAVY":
                case "DASHLONG":
                case "DASHLONGHEAVY":
                    return "DASH";
                case "DOT":
                case "DOTTED":
                case "DOTTEDDASH":
                case "DASHEDHEAVY":
                case "DOTTEDHEAVY":
                    return "DOT";
                default:
                    return "SOLID";
            }
        }

        private static bool TryParseRgbColor(string? colorHex, out double red, out double green, out double blue) {
            red = green = blue = 0;
            if (string.IsNullOrWhiteSpace(colorHex)) {
                return false;
            }

            var normalized = colorHex!.Trim().TrimStart('#');
            if (normalized.Length == 8) {
                normalized = normalized.Substring(2);
            }

            if (normalized.Length != 6) {
                return false;
            }

            if (!int.TryParse(normalized.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var redByte)
                || !int.TryParse(normalized.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var greenByte)
                || !int.TryParse(normalized.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var blueByte)) {
                return false;
            }

            red = redByte / 255d;
            green = greenByte / 255d;
            blue = blueByte / 255d;
            return true;
        }

        private static string SanitizeText(string? value) {
            if (string.IsNullOrEmpty(value)) {
                return string.Empty;
            }

            var normalized = value!.Replace("\r\n", "\n").Replace('\r', '\n');
            var builder = new StringBuilder(normalized.Length);
            foreach (var character in normalized) {
                if (character == '\n' || character == '\t' || !char.IsControl(character)) {
                    builder.Append(character);
                }
            }

            return builder.ToString();
        }

        private static void AddReportNoticeOnce(
            TranslationReport report,
            TranslationSeverity severity,
            string feature,
            string message) {
            if (report.Notices.Any(notice =>
                notice.Severity == severity
                && string.Equals(notice.Feature, feature, StringComparison.Ordinal)
                && string.Equals(notice.Message, message, StringComparison.Ordinal))) {
                return;
            }

            report.Add(severity, feature, message);
        }

        private sealed class MaterializedParagraph {
            public string PrefixText { get; set; } = string.Empty;
            public StringBuilder TextBuilder { get; } = new StringBuilder();
            public List<MaterializedRun> Runs { get; } = new List<MaterializedRun>();
            public List<MaterializedFootnote> Footnotes { get; } = new List<MaterializedFootnote>();
            public List<MaterializedImage> Images { get; } = new List<MaterializedImage>();
            public string InsertedText => PrefixText + TextBuilder.ToString();
        }

        private sealed class MaterializedRun {
            public int StartOffset { get; set; }
            public int EndOffset { get; set; }
            public GoogleDocsParagraphRun Source { get; set; } = new GoogleDocsParagraphRun();
        }

        private sealed class MaterializedImage {
            public int InsertOffset { get; set; }
            public string Uri { get; set; } = string.Empty;
            public GoogleDocsInlineImage Source { get; set; } = new GoogleDocsInlineImage();
        }

        internal sealed class PreparedInitialBatchUpdate {
            public GoogleDocsApiBatchUpdatePayload Payload { get; set; } = new GoogleDocsApiBatchUpdatePayload();
            public List<GoogleDocsFootnote> Footnotes { get; } = new List<GoogleDocsFootnote>();
        }

        internal sealed class PreparedTableContentBatchUpdate {
            public GoogleDocsApiBatchUpdatePayload Payload { get; set; } = new GoogleDocsApiBatchUpdatePayload();
            public List<GoogleDocsFootnote> Footnotes { get; } = new List<GoogleDocsFootnote>();
        }

        private sealed class MaterializedFootnote {
            public int InsertOffset { get; set; }
            public GoogleDocsFootnote Source { get; set; } = new GoogleDocsFootnote();
        }

        private sealed class LiveTableContext {
            public int StartIndex { get; set; }
            public GoogleDocsApiTableResponse Table { get; set; } = new GoogleDocsApiTableResponse();
        }
    }
}

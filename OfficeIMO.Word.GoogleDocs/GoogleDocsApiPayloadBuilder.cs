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
            return BuildInitialBatchUpdatePayload(batch, null);
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildInitialBatchUpdatePayload(
            GoogleDocsBatch batch,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));

            var payload = new GoogleDocsApiBatchUpdatePayload();
            foreach (var request in batch.Requests.Reverse()) {
                AppendRequest(payload, batch.Report, request, imageUris, null);
            }

            return payload;
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
                AppendRequest(payload, report, request, imageUris, segmentId);
            }

            return payload;
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildTableContentBatchUpdatePayload(
            GoogleDocsBatch batch,
            GoogleDocsApiDocumentResponse documentState,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris = null) {
            if (batch == null) throw new ArgumentNullException(nameof(batch));
            if (documentState == null) throw new ArgumentNullException(nameof(documentState));

            var payload = new GoogleDocsApiBatchUpdatePayload();
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
                            imageUris);
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
                            imageUris);
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
            AppendParagraphRequests(payload, report, paragraph, imageUris, segmentId, 1, true);
        }

        private static void AppendParagraphRequests(
            GoogleDocsApiBatchUpdatePayload payload,
            TranslationReport report,
            GoogleDocsParagraph paragraph,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris,
            string? segmentId,
            int insertionIndex,
            bool allowStructuralBreaks) {
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
            string? segmentId) {
            switch (request) {
                case GoogleDocsInsertParagraphRequest paragraphRequest:
                    AppendParagraphRequests(payload, report, paragraphRequest.Paragraph, imageUris, segmentId);
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

        private static List<GoogleDocsApiRequestPayload> BuildCellContentRequests(
            GoogleDocsTableCell cell,
            TranslationReport report,
            int insertionIndex,
            string? segmentId,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris) {
            var payload = new GoogleDocsApiBatchUpdatePayload();
            foreach (var paragraph in cell.Paragraphs.AsEnumerable().Reverse()) {
                AppendParagraphRequests(payload, report, paragraph, imageUris, segmentId, insertionIndex, false);
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

            if (!string.IsNullOrWhiteSpace(source.ForegroundColorHex)) {
                style.ForegroundColor = BuildOptionalColor(source.ForegroundColorHex);
                hasStyle = style.ForegroundColor != null;
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
            if (style.ForegroundColor != null) fields.Add("foregroundColor");
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

        private sealed class LiveTableContext {
            public int StartIndex { get; set; }
            public GoogleDocsApiTableResponse Table { get; set; } = new GoogleDocsApiTableResponse();
        }
    }
}

using OfficeIMO.GoogleWorkspace;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.GoogleDocs {
    internal static partial class GoogleDocsApiPayloadBuilder {
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

        private static bool HasMeaningfulCellContentRequests(IReadOnlyList<GoogleDocsApiRequestPayload> requests) {
            return requests.Any(request => request.InsertText == null
                || !string.Equals(request.InsertText.Text, "\n", StringComparison.Ordinal));
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
    }
}

namespace OfficeIMO.Word.GoogleDocs {
    internal static partial class GoogleDocsApiPayloadBuilder {
        internal static void ApplyTabId(GoogleDocsApiBatchUpdatePayload payload, string? tabId) {
            if (payload == null) throw new ArgumentNullException(nameof(payload));
            if (string.IsNullOrWhiteSpace(tabId)) return;
            string selectedTabId = tabId!;
            foreach (GoogleDocsApiRequestPayload request in payload.Requests) {
                Set(request.InsertText?.Location, selectedTabId);
                Set(request.InsertText?.EndOfSegmentLocation, selectedTabId);
                Set(request.CreateHeader?.SectionBreakLocation, selectedTabId);
                Set(request.CreateFooter?.SectionBreakLocation, selectedTabId);
                Set(request.CreateFootnote?.Location, selectedTabId);
                Set(request.CreateNamedRange?.Range, selectedTabId);
                Set(request.InsertInlineImage?.Location, selectedTabId);
                Set(request.UpdateTextStyle?.Range, selectedTabId);
                Set(request.UpdateParagraphStyle?.Range, selectedTabId);
                Set(request.UpdateSectionStyle?.Range, selectedTabId);
                Set(request.CreateParagraphBullets?.Range, selectedTabId);
                Set(request.InsertTable?.Location, selectedTabId);
                Set(request.MergeTableCells?.TableRange.TableCellLocation.TableStartLocation, selectedTabId);
                Set(request.PinTableHeaderRows?.TableStartLocation, selectedTabId);
                Set(request.UpdateTableCellStyle?.TableRange.TableCellLocation.TableStartLocation, selectedTabId);
                Set(request.UpdateTableColumnProperties?.TableStartLocation, selectedTabId);
                Set(request.InsertPageBreak?.Location, selectedTabId);
                Set(request.InsertSectionBreak?.Location, selectedTabId);
                Set(request.DeleteContentRange?.Range, selectedTabId);
            }
        }

        internal static GoogleDocsApiBatchUpdatePayload BuildResetDocumentPayload(
            GoogleDocsApiDocumentResponse documentState,
            GoogleDocsTabOptions options) {
            if (documentState == null) throw new ArgumentNullException(nameof(documentState));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (documentState.Tabs.Count == 0) return BuildResetDocumentPayload(documentState);

            var payload = new GoogleDocsApiBatchUpdatePayload();
            IReadOnlyList<GoogleDocsApiTabResponse> selected = SelectTabs(documentState, options);
            foreach (GoogleDocsApiTabResponse tab in selected) {
                string? tabId = tab.Properties.TabId;
                int endIndex = GetBodyEndIndex(tab.DocumentTab?.Body);
                if (endIndex > 2) {
                    payload.Requests.Add(new GoogleDocsApiRequestPayload {
                        DeleteContentRange = new GoogleDocsApiDeleteContentRangeRequestPayload {
                            Range = new GoogleDocsApiRangePayload { StartIndex = 1, EndIndex = endIndex - 1, TabId = tabId }
                        }
                    });
                }
                foreach (string id in tab.DocumentTab?.Headers?.Keys ?? Enumerable.Empty<string>()) {
                    payload.Requests.Add(new GoogleDocsApiRequestPayload { DeleteHeader = new GoogleDocsApiDeleteHeaderRequestPayload { HeaderId = id, TabId = tabId } });
                }
                foreach (string id in tab.DocumentTab?.Footers?.Keys ?? Enumerable.Empty<string>()) {
                    payload.Requests.Add(new GoogleDocsApiRequestPayload { DeleteFooter = new GoogleDocsApiDeleteFooterRequestPayload { FooterId = id, TabId = tabId } });
                }
                foreach (string name in tab.DocumentTab?.NamedRanges?.Keys ?? Enumerable.Empty<string>()) {
                    payload.Requests.Add(new GoogleDocsApiRequestPayload {
                        DeleteNamedRange = new GoogleDocsApiDeleteNamedRangeRequestPayload {
                            Name = name,
                            TabsCriteria = string.IsNullOrWhiteSpace(tabId) ? null : new GoogleDocsApiTabsCriteriaPayload { TabIds = new List<string> { tabId! } },
                        }
                    });
                }
            }
            return payload;
        }

        internal static IReadOnlyList<GoogleDocsApiTabResponse> SelectTabs(GoogleDocsApiDocumentResponse document, GoogleDocsTabOptions options) {
            var all = FlattenTabs(document.Tabs).ToList();
            if (all.Count == 0) return Array.Empty<GoogleDocsApiTabResponse>();
            return options.Strategy switch {
                GoogleDocsTabStrategy.ReplaceEveryTab => all,
                GoogleDocsTabStrategy.SelectedTab => new[] { all.FirstOrDefault(tab => string.Equals(tab.Properties.TabId, options.TabId, StringComparison.Ordinal))
                    ?? throw new InvalidOperationException($"Google document does not contain tab '{options.TabId}'.") },
                _ => new[] { all[0] },
            };
        }

        internal static IEnumerable<GoogleDocsApiTabResponse> FlattenTabs(IEnumerable<GoogleDocsApiTabResponse> tabs) {
            foreach (GoogleDocsApiTabResponse tab in tabs) {
                yield return tab;
                foreach (GoogleDocsApiTabResponse child in FlattenTabs(tab.ChildTabs)) yield return child;
            }
        }

        private static int GetBodyEndIndex(GoogleDocsApiBodyResponse? body) {
            return body?.Content?.Where(element => element.EndIndex.HasValue).Select(element => element.EndIndex!.Value).DefaultIfEmpty(1).Max() ?? 1;
        }

        private static void Set(GoogleDocsApiLocationPayload? location, string tabId) { if (location != null) location.TabId = tabId; }
        private static void Set(GoogleDocsApiRangePayload? range, string tabId) { if (range != null) range.TabId = tabId; }
        private static void Set(GoogleDocsApiEndOfSegmentLocationPayload? location, string tabId) { if (location != null) location.TabId = tabId; }
    }
}

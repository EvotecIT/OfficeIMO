using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Word.GoogleDocs {
    internal sealed class GoogleDocsWriteControlState {
        private readonly GoogleDocsRevisionConflictMode _mode;
        private string? _revisionId;
        private bool _writeAttempted;

        internal GoogleDocsWriteControlState(GoogleDocsRevisionConflictMode mode, string? revisionId) {
            _mode = mode;
            _revisionId = revisionId;
        }

        internal string? RevisionId => _revisionId;
        internal bool RequiresRevisionRefresh => _mode == GoogleDocsRevisionConflictMode.OverwriteLatest && _writeAttempted;

        internal void Apply(GoogleDocsApiBatchUpdatePayload payload) {
            if (_mode == GoogleDocsRevisionConflictMode.OverwriteLatest) {
                _writeAttempted = true;
                return;
            }
            if (string.IsNullOrWhiteSpace(_revisionId)) return;
            payload.WriteControl = _mode == GoogleDocsRevisionConflictMode.RequireRevision
                ? new GoogleDocsApiWriteControlPayload { RequiredRevisionId = _revisionId }
                : new GoogleDocsApiWriteControlPayload { TargetRevisionId = _revisionId };
        }

        internal void Observe(GoogleDocsApiBatchUpdateResponse response) {
            string? updated = response.WriteControl?.RequiredRevisionId ?? response.WriteControl?.TargetRevisionId;
            if (!string.IsNullOrWhiteSpace(updated)) {
                _revisionId = updated;
            }
        }
    }

    public sealed partial class GoogleDocsExporter {
        private static async Task<GoogleDocsApiDocumentResponse> GetDocumentAsync(
            GoogleWorkspaceHttpTransport transport,
            string accessToken,
            string documentId,
            GoogleDocsBatch batch,
            CancellationToken cancellationToken) {
            GoogleDocsApiDocumentResponse document = await transport.SendJsonAsync<GoogleDocsApiDocumentResponse>(
                accessToken,
                HttpMethod.Get,
                $"https://docs.googleapis.com/v1/documents/{documentId}?includeTabsContent=true",
                null,
                GoogleWorkspaceRequestSafety.Safe,
                "Google Docs API",
                batch.Report,
                GoogleDocsJsonSerializerContext.Default.GoogleDocsApiDocumentResponse,
                cancellationToken).ConfigureAwait(false);
            ProjectSelectedTab(document, batch.TargetTabId);
            return document;
        }

        private static void ProjectSelectedTab(GoogleDocsApiDocumentResponse document, string? tabId) {
            if (document.Tabs.Count == 0) return;
            GoogleDocsApiTabResponse? tab = GoogleDocsApiPayloadBuilder.FlattenTabs(document.Tabs)
                .FirstOrDefault(candidate => string.IsNullOrWhiteSpace(tabId) || string.Equals(candidate.Properties.TabId, tabId, StringComparison.Ordinal));
            if (tab?.DocumentTab == null) return;
            document.Body = tab.DocumentTab.Body;
            document.Headers = tab.DocumentTab.Headers;
            document.Footers = tab.DocumentTab.Footers;
            document.Footnotes = tab.DocumentTab.Footnotes;
        }

        private static async Task<GoogleDocsApiBatchUpdateResponse> SendBatchUpdateAsync(
            GoogleWorkspaceHttpTransport transport,
            string accessToken,
            string documentId,
            GoogleDocsBatch batch,
            GoogleDocsApiBatchUpdatePayload payload,
            CancellationToken cancellationToken) {
            GoogleDocsApiPayloadBuilder.ApplyTabId(payload, batch.TargetTabId);
            batch.WriteControlState?.Apply(payload);
            GoogleDocsApiBatchUpdateResponse response = await transport.SendJsonAsync<GoogleDocsApiBatchUpdatePayload, GoogleDocsApiBatchUpdateResponse>(
                accessToken,
                HttpMethod.Post,
                $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                payload,
                GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Docs API",
                batch.Report,
                GoogleDocsJsonSerializerContext.Default.GoogleDocsApiBatchUpdatePayload,
                GoogleDocsJsonSerializerContext.Default.GoogleDocsApiBatchUpdateResponse,
                cancellationToken).ConfigureAwait(false);
            batch.WriteControlState?.Observe(response);
            return response;
        }
    }
}

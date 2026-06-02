using OfficeIMO.GoogleWorkspace;
using System.Net.Http.Headers;
using System.IO;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Word.GoogleDocs {
    public sealed partial class GoogleDocsExporter : IGoogleDocsExporter {
        private static async Task<GoogleDriveFileMetadataResponse?> ApplyDrivePlacementAsync(
            HttpClient client,
            string accessToken,
            string? fileId,
            GoogleDriveFileLocation location,
            GoogleWorkspaceRetryOptions retryOptions,
            TranslationReport report,
            CancellationToken cancellationToken) {
            if (string.IsNullOrWhiteSpace(fileId) || string.IsNullOrWhiteSpace(location.FolderId)) {
                return null;
            }

            var supportsAllDrives = location.SharedDriveAware || !string.IsNullOrWhiteSpace(location.DriveId);
            var supportsAllDrivesQuery = supportsAllDrives ? "&supportsAllDrives=true" : string.Empty;
            var currentMetadata = await SendAsync<GoogleDriveFileMetadataResponse>(
                client,
                accessToken,
                HttpMethod.Get,
                $"https://www.googleapis.com/drive/v3/files/{fileId}?fields=id,parents,webViewLink{supportsAllDrivesQuery}",
                null,
                retryOptions,
                report,
                cancellationToken).ConfigureAwait(false);

            var desiredFolderId = location.FolderId!;
            if (currentMetadata.Parents.Count == 1 && string.Equals(currentMetadata.Parents[0], desiredFolderId, StringComparison.OrdinalIgnoreCase)) {
                return currentMetadata;
            }

            var query = new List<string> {
                "supportsAllDrives=" + (supportsAllDrives ? "true" : "false"),
                "addParents=" + Uri.EscapeDataString(desiredFolderId),
                "fields=id,parents,webViewLink"
            };

            if (currentMetadata.Parents.Count > 0) {
                query.Add("removeParents=" + Uri.EscapeDataString(string.Join(",", currentMetadata.Parents)));
            }

            return await SendAsync<GoogleDriveFileMetadataResponse>(
                client,
                accessToken,
                new HttpMethod("PATCH"),
                $"https://www.googleapis.com/drive/v3/files/{fileId}?{string.Join("&", query)}",
                new { },
                retryOptions,
                report,
                cancellationToken).ConfigureAwait(false);
        }


        private static async Task<TResponse> SendAsync<TResponse>(
            HttpClient client,
            string accessToken,
            HttpMethod method,
            string uri,
            object? payload,
            GoogleWorkspaceRetryOptions retryOptions,
            TranslationReport report,
            CancellationToken cancellationToken) {
            return await SendAsync<TResponse>(
                client,
                accessToken,
                method,
                uri,
                payload == null ? null : (() => payload),
                retryOptions,
                report,
                cancellationToken).ConfigureAwait(false);
        }

        private static async Task<TResponse> SendAsync<TResponse>(
            HttpClient client,
            string accessToken,
            HttpMethod method,
            string uri,
            Func<object?>? payloadFactory,
            GoogleWorkspaceRetryOptions retryOptions,
            TranslationReport report,
            CancellationToken cancellationToken) {
            using (var response = await GoogleWorkspaceRetryPolicy.SendAsync(
                client,
                () => {
                    var request = new HttpRequestMessage(method, uri);
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    if (payloadFactory != null) {
                        var payload = payloadFactory();
                        if (payload is HttpContent httpContent) {
                            request.Content = httpContent;
                        } else if (payload != null) {
                            var json = JsonSerializer.Serialize(payload, JsonOptions);
                            request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                        }
                    }

                    return request;
                },
                retryOptions,
                cancellationToken,
                retryEvent => ReportRetry(report, retryOptions.SessionOptions, "Google Docs API", retryEvent)).ConfigureAwait(false)) {
                var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode) {
                    string formattedError = GoogleWorkspaceApiErrorFormatter.Format(body) ?? body;
                    throw new HttpRequestException($"Google Docs API request to '{uri}' failed with {(int)response.StatusCode}: {formattedError}");
                }

                if (typeof(TResponse) == typeof(object) || string.IsNullOrWhiteSpace(body)) {
                    return default!;
                }

                var result = JsonSerializer.Deserialize<TResponse>(body, JsonOptions);
                if (result == null) {
                    throw new InvalidOperationException($"Google Docs API response from '{uri}' could not be deserialized.");
                }

                return result;
            }
        }

        private static void ReportRetry(TranslationReport report, GoogleWorkspaceSessionOptions? sessionOptions, string serviceName, GoogleWorkspaceRetryEvent retryEvent) {
            GoogleWorkspaceDiagnosticsDispatcher.AddUnique(
                report,
                sessionOptions,
                TranslationSeverity.Info,
                "ApiRetries",
                $"{serviceName} retried {retryEvent.Method} {retryEvent.Uri} after transient {retryEvent.Trigger} using {retryEvent.DelayStrategy} ({retryEvent.Delay.TotalMilliseconds:0} ms, retry {retryEvent.RetryAttempt} of {retryEvent.MaxRetryCount}).",
                $"{retryEvent.Method} {retryEvent.Uri}");
        }

        private static string? BuildDocumentWebViewLink(string? documentId) {
            return string.IsNullOrWhiteSpace(documentId)
                ? null
                : $"https://docs.google.com/document/d/{documentId}/edit";
        }
    }

    internal sealed class GoogleDriveFileMetadataResponse {
        [System.Text.Json.Serialization.JsonPropertyName("id")]
        public string? Id { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("parents")]
        public List<string> Parents { get; set; } = new List<string>();

        [System.Text.Json.Serialization.JsonPropertyName("webViewLink")]
        public string? WebViewLink { get; set; }
    }
}

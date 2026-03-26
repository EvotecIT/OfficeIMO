using OfficeIMO.GoogleWorkspace;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Default Excel to Google Sheets exporter implementation.
    /// </summary>
    public sealed class GoogleSheetsExporter : IGoogleSheetsExporter {
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions {
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
            PropertyNamingPolicy = null,
            WriteIndented = false,
        };

        public GoogleSheetsTranslationPlan BuildPlan(ExcelDocument document, GoogleSheetsSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return GoogleSheetsPlanBuilder.Build(document, options ?? new GoogleSheetsSaveOptions());
        }

        public GoogleSheetsBatch BuildBatch(ExcelDocument document, GoogleSheetsSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return GoogleSheetsBatchCompiler.Build(document, options ?? new GoogleSheetsSaveOptions());
        }

        public async Task<GoogleSpreadsheetReference> ExportAsync(
            ExcelDocument document,
            GoogleWorkspaceSession session,
            GoogleSheetsSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (session == null) throw new ArgumentNullException(nameof(session));

            var effectiveOptions = options ?? new GoogleSheetsSaveOptions();
            var batch = BuildBatch(document, effectiveOptions);
            var effectiveLocation = session.ResolveLocationDefaults(effectiveOptions.Location);
            if (string.IsNullOrWhiteSpace(effectiveLocation.FolderId) && !string.IsNullOrWhiteSpace(effectiveLocation.DriveId)) {
                GoogleWorkspaceDiagnosticsDispatcher.Add(
                    batch.Report,
                    session.Options,
                    TranslationSeverity.Warning,
                    "DrivePlacement",
                    "Drive placement requires a concrete FolderId. Supplying DriveId without FolderId is still treated as diagnostic-only.");
            }

            GoogleWorkspaceAccessToken accessToken;
            try {
                accessToken = await session.AcquireAccessTokenAsync(GoogleWorkspaceScopeCatalog.SheetsAuthoring, cancellationToken).ConfigureAwait(false);
            } catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested) {
                throw GoogleWorkspaceFailureDiagnostics.CreateRequestTimeoutFailure(
                    "Google Sheets export token acquisition",
                    session.Options,
                    batch.Report,
                    ex);
            } catch (OperationCanceledException ex) when (cancellationToken.IsCancellationRequested) {
                throw GoogleWorkspaceFailureDiagnostics.CreateCanceledFailure(
                    "Google Sheets export",
                    session.Options,
                    batch.Report,
                    ex);
            } catch (Exception ex) when (!(ex is OperationCanceledException)) {
                throw GoogleWorkspaceFailureDiagnostics.CreateTokenAcquisitionFailure(
                    "Google Sheets export",
                    GoogleWorkspaceScopeCatalog.SheetsAuthoring,
                    session,
                    batch.Report,
                    ex);
            }

            var retryOptions = GoogleWorkspaceRetryOptions.FromSessionOptions(session.Options);

            bool disposeClient = session.Options.HttpClient == null;
            var client = session.Options.HttpClient ?? new HttpClient();
            try {
                client.Timeout = session.Options.RequestTimeout;

                if (!string.IsNullOrWhiteSpace(effectiveLocation.ExistingFileId)) {
                    var existingResponse = await SendAsync<GoogleSheetsApiSpreadsheetMetadataResponse>(
                        client,
                        accessToken.AccessToken,
                        HttpMethod.Get,
                        $"https://sheets.googleapis.com/v4/spreadsheets/{effectiveLocation.ExistingFileId}?fields=spreadsheetId,spreadsheetUrl,properties.title,sheets.properties.sheetId",
                        null,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);

                    var existingSheetIds = existingResponse.Sheets
                        .Select(sheet => sheet.Properties?.SheetId ?? 0)
                        .Where(sheetId => sheetId > 0)
                        .ToList();
                    var sheetIdMap = GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch, existingSheetIds);
                    var replacePayload = GoogleSheetsApiPayloadBuilder.BuildReplaceSpreadsheetPayload(batch, existingSheetIds, sheetIdMap);

                    await SendAsync<object>(
                        client,
                        accessToken.AccessToken,
                        HttpMethod.Post,
                        $"https://sheets.googleapis.com/v4/spreadsheets/{effectiveLocation.ExistingFileId}:batchUpdate",
                        replacePayload,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);

                    var contentPayload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(
                        batch,
                        sheetIdMap,
                        existingResponse.SpreadsheetId ?? effectiveLocation.ExistingFileId);
                    if (contentPayload.Requests.Count > 0) {
                        await SendAsync<object>(
                            client,
                            accessToken.AccessToken,
                            HttpMethod.Post,
                            $"https://sheets.googleapis.com/v4/spreadsheets/{effectiveLocation.ExistingFileId}:batchUpdate",
                            contentPayload,
                            retryOptions,
                            batch.Report,
                            cancellationToken).ConfigureAwait(false);
                    }

                    var updatedDriveMetadata = await ApplyDrivePlacementAsync(
                        client,
                        accessToken.AccessToken,
                        existingResponse.SpreadsheetId ?? effectiveLocation.ExistingFileId!,
                        effectiveLocation,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);

                    batch.Report.Add(
                        TranslationSeverity.Info,
                        "ExistingSpreadsheet",
                        "Existing spreadsheet replacement currently recreates workbook sheets before replaying the OfficeIMO batch.");

                    return new GoogleSpreadsheetReference {
                        SpreadsheetId = existingResponse.SpreadsheetId ?? effectiveLocation.ExistingFileId,
                        FileId = existingResponse.SpreadsheetId ?? effectiveLocation.ExistingFileId,
                        Name = batch.Title,
                        MimeType = "application/vnd.google-apps.spreadsheet",
                        WebViewLink = updatedDriveMetadata?.WebViewLink
                            ?? (!string.IsNullOrWhiteSpace(existingResponse.SpreadsheetUrl)
                                ? existingResponse.SpreadsheetUrl
                                : BuildSpreadsheetWebViewLink(existingResponse.SpreadsheetId ?? effectiveLocation.ExistingFileId)),
                        Location = effectiveLocation,
                        Report = batch.Report,
                    };
                }

                var sheetIdMapForCreate = GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch);
                var createPayload = GoogleSheetsApiPayloadBuilder.BuildCreateSpreadsheetPayload(batch, sheetIdMapForCreate);
                var createResponse = await SendAsync<GoogleSheetsApiCreateSpreadsheetResponse>(
                    client,
                    accessToken.AccessToken,
                    HttpMethod.Post,
                    "https://sheets.googleapis.com/v4/spreadsheets",
                    createPayload,
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);

                var updatePayload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(
                    batch,
                    sheetIdMapForCreate,
                    createResponse.SpreadsheetId);

                if (!string.IsNullOrWhiteSpace(createResponse.SpreadsheetId) && updatePayload.Requests.Count > 0) {
                    await SendAsync<object>(
                        client,
                        accessToken.AccessToken,
                        HttpMethod.Post,
                        $"https://sheets.googleapis.com/v4/spreadsheets/{createResponse.SpreadsheetId}:batchUpdate",
                        updatePayload,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);
                }

                var createdDriveMetadata = await ApplyDrivePlacementAsync(
                    client,
                    accessToken.AccessToken,
                    createResponse.SpreadsheetId,
                    effectiveLocation,
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);

                return new GoogleSpreadsheetReference {
                    SpreadsheetId = createResponse.SpreadsheetId,
                    FileId = createResponse.SpreadsheetId,
                    Name = createResponse.Properties?.Title ?? batch.Title,
                    MimeType = "application/vnd.google-apps.spreadsheet",
                    WebViewLink = createdDriveMetadata?.WebViewLink
                        ?? (!string.IsNullOrWhiteSpace(createResponse.SpreadsheetUrl)
                            ? createResponse.SpreadsheetUrl
                            : BuildSpreadsheetWebViewLink(createResponse.SpreadsheetId)),
                    Location = effectiveLocation,
                    Report = batch.Report,
                };
            } catch (GoogleWorkspaceExportException) {
                throw;
            } catch (GoogleWorkspaceExportCanceledException) {
                throw;
            } catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested) {
                throw GoogleWorkspaceFailureDiagnostics.CreateRequestTimeoutFailure(
                    "Google Sheets export",
                    session.Options,
                    batch.Report,
                    ex);
            } catch (OperationCanceledException ex) when (cancellationToken.IsCancellationRequested) {
                throw GoogleWorkspaceFailureDiagnostics.CreateCanceledFailure(
                    "Google Sheets export",
                    session.Options,
                    batch.Report,
                    ex);
            } catch (Exception ex) when (!(ex is OperationCanceledException)) {
                throw GoogleWorkspaceFailureDiagnostics.CreateApiFailure(
                    "Google Sheets export",
                    session.Options,
                    batch.Report,
                    ex);
            } finally {
                if (disposeClient) {
                    client.Dispose();
                }
            }
        }

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
            using (var response = await GoogleWorkspaceRetryPolicy.SendAsync(
                client,
                () => {
                    var request = new HttpRequestMessage(method, uri);
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    if (payload != null) {
                        var json = JsonSerializer.Serialize(payload, JsonOptions);
                        request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                    }

                    return request;
                },
                retryOptions,
                cancellationToken,
                retryEvent => ReportRetry(report, retryOptions.SessionOptions, "Google Sheets API", retryEvent)).ConfigureAwait(false)) {
                var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode) {
                    string formattedError = GoogleWorkspaceApiErrorFormatter.Format(body) ?? body;
                    throw new HttpRequestException($"Google Sheets API request to '{uri}' failed with {(int)response.StatusCode}: {formattedError}");
                }

                if (typeof(TResponse) == typeof(object) || string.IsNullOrWhiteSpace(body)) {
                    return default!;
                }

                var result = JsonSerializer.Deserialize<TResponse>(body, JsonOptions);
                if (result == null) {
                    throw new InvalidOperationException($"Google Sheets API response from '{uri}' could not be deserialized.");
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

        private static string? BuildSpreadsheetWebViewLink(string? spreadsheetId) {
            return string.IsNullOrWhiteSpace(spreadsheetId)
                ? null
                : $"https://docs.google.com/spreadsheets/d/{spreadsheetId}/edit";
        }
    }

    internal sealed class GoogleSheetsApiCreateSpreadsheetResponse {
        [System.Text.Json.Serialization.JsonPropertyName("spreadsheetId")]
        public string? SpreadsheetId { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("spreadsheetUrl")]
        public string? SpreadsheetUrl { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("properties")]
        public GoogleSheetsApiSpreadsheetPropertiesPayload? Properties { get; set; }
    }

    internal sealed class GoogleSheetsApiSpreadsheetMetadataResponse {
        [System.Text.Json.Serialization.JsonPropertyName("spreadsheetId")]
        public string? SpreadsheetId { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("spreadsheetUrl")]
        public string? SpreadsheetUrl { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("properties")]
        public GoogleSheetsApiSpreadsheetPropertiesPayload? Properties { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("sheets")]
        public List<GoogleSheetsApiSheetMetadataResponse> Sheets { get; set; } = new List<GoogleSheetsApiSheetMetadataResponse>();
    }

    internal sealed class GoogleSheetsApiSheetMetadataResponse {
        [System.Text.Json.Serialization.JsonPropertyName("properties")]
        public GoogleSheetsApiSheetMetadataPropertiesResponse? Properties { get; set; }
    }

    internal sealed class GoogleSheetsApiSheetMetadataPropertiesResponse {
        [System.Text.Json.Serialization.JsonPropertyName("sheetId")]
        public int SheetId { get; set; }
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

using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Default Excel to Google Sheets exporter implementation.
    /// </summary>
    public sealed class GoogleSheetsExporter : IGoogleSheetsExporter {
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
            GoogleWorkspacePreflight.Validate(batch.Report, effectiveOptions.FidelityPolicy);
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

            using (var transport = new GoogleWorkspaceHttpTransport(session.Options)) {
            using (var driveClient = new GoogleDriveClient(session)) {
            try {

                if (!string.IsNullOrWhiteSpace(effectiveLocation.ExistingFileId)) {
                    var existingResponse = await transport.SendJsonAsync<GoogleSheetsApiSpreadsheetMetadataResponse>(
                        accessToken.AccessToken,
                        HttpMethod.Get,
                        $"https://sheets.googleapis.com/v4/spreadsheets/{effectiveLocation.ExistingFileId}?fields=spreadsheetId,spreadsheetUrl,properties.title,sheets.properties.sheetId",
                        null,
                        GoogleWorkspaceRequestSafety.Safe,
                        "Google Sheets API",
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);

                    var existingSheetIds = existingResponse.Sheets
                        .Select(sheet => sheet.Properties?.SheetId ?? 0)
                        .Where(sheetId => sheetId > 0)
                        .ToList();
                    var sheetIdMap = GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch, existingSheetIds);
                    var replacePayload = GoogleSheetsApiPayloadBuilder.BuildReplaceSpreadsheetPayload(batch, existingSheetIds, sheetIdMap);

                    await transport.SendJsonAsync<object>(
                        accessToken.AccessToken,
                        HttpMethod.Post,
                        $"https://sheets.googleapis.com/v4/spreadsheets/{effectiveLocation.ExistingFileId}:batchUpdate",
                        replacePayload,
                        GoogleWorkspaceRequestSafety.NonIdempotent,
                        "Google Sheets API",
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);

                    var contentPayload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(
                        batch,
                        sheetIdMap,
                        existingResponse.SpreadsheetId ?? effectiveLocation.ExistingFileId);
                    if (contentPayload.Requests.Count > 0) {
                        await transport.SendJsonAsync<object>(
                            accessToken.AccessToken,
                            HttpMethod.Post,
                            $"https://sheets.googleapis.com/v4/spreadsheets/{effectiveLocation.ExistingFileId}:batchUpdate",
                            contentPayload,
                            GoogleWorkspaceRequestSafety.NonIdempotent,
                            "Google Sheets API",
                            batch.Report,
                            cancellationToken).ConfigureAwait(false);
                    }

                    var updatedDriveMetadata = await ApplyDrivePlacementAsync(
                        driveClient,
                        existingResponse.SpreadsheetId ?? effectiveLocation.ExistingFileId!,
                        effectiveLocation,
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
                var createResponse = await transport.SendJsonAsync<GoogleSheetsApiCreateSpreadsheetResponse>(
                    accessToken.AccessToken,
                    HttpMethod.Post,
                    "https://sheets.googleapis.com/v4/spreadsheets",
                    createPayload,
                    GoogleWorkspaceRequestSafety.NonIdempotent,
                    "Google Sheets API",
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);

                var updatePayload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(
                    batch,
                    sheetIdMapForCreate,
                    createResponse.SpreadsheetId);

                if (!string.IsNullOrWhiteSpace(createResponse.SpreadsheetId) && updatePayload.Requests.Count > 0) {
                    await transport.SendJsonAsync<object>(
                        accessToken.AccessToken,
                        HttpMethod.Post,
                        $"https://sheets.googleapis.com/v4/spreadsheets/{createResponse.SpreadsheetId}:batchUpdate",
                        updatePayload,
                        GoogleWorkspaceRequestSafety.NonIdempotent,
                        "Google Sheets API",
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);
                }

                var createdDriveMetadata = await ApplyDrivePlacementAsync(
                    driveClient,
                    createResponse.SpreadsheetId,
                    effectiveLocation,
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
            }
            }
            }
        }

        private static async Task<GoogleDriveFile?> ApplyDrivePlacementAsync(
            GoogleDriveClient driveClient,
            string? fileId,
            GoogleDriveFileLocation location,
            TranslationReport report,
            CancellationToken cancellationToken) {
            if (string.IsNullOrWhiteSpace(fileId) || string.IsNullOrWhiteSpace(location.FolderId)) {
                return null;
            }

            await driveClient.ResolveFolderAsync(location.FolderId!, location.DriveId, report, cancellationToken).ConfigureAwait(false);
            return await driveClient.MoveFileAsync(fileId!, location.FolderId!, report, cancellationToken).ConfigureAwait(false);
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

}

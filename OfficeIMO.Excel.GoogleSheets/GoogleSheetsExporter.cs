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
            ValidateExecutionOptions(effectiveOptions.Execution);
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
            using (var driveClient = new GoogleDriveClient(session, GoogleDriveClientOptions.ForFileAuthoring())) {
            try {
                await ValidateDrivePlacementAsync(
                    driveClient,
                    effectiveLocation,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);

                if (!string.IsNullOrWhiteSpace(effectiveLocation.ExistingFileId)) {
                    await ValidateReplaceTargetAsync(
                        driveClient,
                        effectiveLocation.ExistingFileId!,
                        effectiveOptions.Replace,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);

                    var existingResponse = await transport.SendJsonAsync<GoogleSheetsApiSpreadsheetMetadataResponse>(
                        accessToken.AccessToken,
                        HttpMethod.Get,
                        $"https://sheets.googleapis.com/v4/spreadsheets/{effectiveLocation.ExistingFileId}?fields=spreadsheetId,spreadsheetUrl,properties.title,sheets(properties(sheetId,title))",
                        null,
                        GoogleWorkspaceRequestSafety.Safe,
                        "Google Sheets API",
                        batch.Report,
                        GoogleSheetsJsonSerializerContext.Default.GoogleSheetsApiSpreadsheetMetadataResponse,
                        cancellationToken).ConfigureAwait(false);

                    var existingSheets = existingResponse.Sheets
                        .Where(sheet => sheet.Properties?.SheetId.HasValue == true)
                        .ToDictionary(
                            sheet => sheet.Properties!.SheetId.GetValueOrDefault(),
                            sheet => sheet.Properties!.Title ?? string.Empty);
                    var sheetIdMap = GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch, existingSheets.Keys);
                    var replacePayload = GoogleSheetsApiPayloadBuilder.BuildReplaceSpreadsheetPayload(batch, existingSheets, sheetIdMap);

                    await transport.SendJsonAsync<GoogleSheetsApiBatchUpdatePayload, object>(
                        accessToken.AccessToken,
                        HttpMethod.Post,
                        $"https://sheets.googleapis.com/v4/spreadsheets/{effectiveLocation.ExistingFileId}:batchUpdate",
                        replacePayload,
                        GoogleWorkspaceRequestSafety.NonIdempotent,
                        "Google Sheets API",
                        batch.Report,
                        GoogleSheetsJsonSerializerContext.Default.GoogleSheetsApiBatchUpdatePayload,
                        GoogleSheetsJsonSerializerContext.Default.Object,
                        cancellationToken).ConfigureAwait(false);

                    await SendValuesAsync(
                        transport,
                        accessToken.AccessToken,
                        existingResponse.SpreadsheetId ?? effectiveLocation.ExistingFileId!,
                        batch,
                        sheetIdMap,
                        effectiveOptions.Execution,
                        cancellationToken).ConfigureAwait(false);

                    var contentPayload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(
                        batch,
                        sheetIdMap,
                        existingResponse.SpreadsheetId ?? effectiveLocation.ExistingFileId,
                        includeCellValues: !effectiveOptions.Execution.UseValuesBatchUpdate);
                    await SendStructuralBatchesAsync(
                        transport,
                        accessToken.AccessToken,
                        effectiveLocation.ExistingFileId!,
                        contentPayload,
                        effectiveOptions.Execution,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);

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
                        DriveVersion = updatedDriveMetadata?.Version,
                        ModifiedTime = updatedDriveMetadata?.ModifiedTime,
                        Location = effectiveLocation,
                        Report = batch.Report,
                    };
                }

                var sheetIdMapForCreate = GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch);
                var createPayload = GoogleSheetsApiPayloadBuilder.BuildCreateSpreadsheetPayload(batch, sheetIdMapForCreate);
                var createResponse = await transport.SendJsonAsync<GoogleSheetsApiCreateSpreadsheetPayload, GoogleSheetsApiCreateSpreadsheetResponse>(
                    accessToken.AccessToken,
                    HttpMethod.Post,
                    "https://sheets.googleapis.com/v4/spreadsheets",
                    createPayload,
                    GoogleWorkspaceRequestSafety.NonIdempotent,
                    "Google Sheets API",
                    batch.Report,
                    GoogleSheetsJsonSerializerContext.Default.GoogleSheetsApiCreateSpreadsheetPayload,
                    GoogleSheetsJsonSerializerContext.Default.GoogleSheetsApiCreateSpreadsheetResponse,
                    cancellationToken).ConfigureAwait(false);

                await SendValuesAsync(
                    transport,
                    accessToken.AccessToken,
                    createResponse.SpreadsheetId!,
                    batch,
                    sheetIdMapForCreate,
                    effectiveOptions.Execution,
                    cancellationToken).ConfigureAwait(false);

                var updatePayload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(
                    batch,
                    sheetIdMapForCreate,
                    createResponse.SpreadsheetId,
                    includeCellValues: !effectiveOptions.Execution.UseValuesBatchUpdate);

                if (!string.IsNullOrWhiteSpace(createResponse.SpreadsheetId) && updatePayload.Requests.Count > 0) {
                    await SendStructuralBatchesAsync(
                        transport,
                        accessToken.AccessToken,
                        createResponse.SpreadsheetId!,
                        updatePayload,
                        effectiveOptions.Execution,
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
                    DriveVersion = createdDriveMetadata?.Version,
                    ModifiedTime = createdDriveMetadata?.ModifiedTime,
                    Location = effectiveLocation,
                    Report = batch.Report,
                };
            } catch (GoogleWorkspaceExportException) {
                throw;
            } catch (GoogleWorkspaceConflictException) {
                throw;
            } catch (GoogleWorkspacePreflightException) {
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
            if (string.IsNullOrWhiteSpace(fileId)) {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(location.FolderId)) {
                return await driveClient.MoveFileAsync(fileId!, location.FolderId!, report, cancellationToken).ConfigureAwait(false);
            }

            return await driveClient.GetFileAsync(fileId!, report: report, cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        private static async Task ValidateDrivePlacementAsync(
            GoogleDriveClient driveClient,
            GoogleDriveFileLocation location,
            TranslationReport report,
            CancellationToken cancellationToken) {
            if (string.IsNullOrWhiteSpace(location.FolderId)) return;
            await driveClient.ResolveFolderAsync(location.FolderId!, location.DriveId, report, cancellationToken).ConfigureAwait(false);
        }

        private static async Task ValidateReplaceTargetAsync(
            GoogleDriveClient driveClient,
            string fileId,
            GoogleSheetsReplaceOptions replace,
            TranslationReport report,
            CancellationToken cancellationToken) {
            if (replace == null) throw new ArgumentNullException(nameof(replace));
            if (replace.ConflictMode == GoogleSheetsReplaceConflictMode.RequireMatchingDriveVersion
                && !replace.ExpectedDriveVersion.HasValue) {
                report.Add(
                    TranslationSeverity.Error,
                    "ReplaceConflict",
                    "Destructive spreadsheet replacement requires the Drive version observed by a prior read/import.",
                    code: "SHEETS.REPLACE.EXPECTED_VERSION_REQUIRED",
                    action: TranslationAction.Fail,
                    targetId: fileId);
                throw new GoogleWorkspacePreflightException(
                    "Google Sheets replacement requires Replace.ExpectedDriveVersion unless ConflictMode is explicitly Overwrite.",
                    report,
                    report.Notices.Where(notice => notice.Code == "SHEETS.REPLACE.EXPECTED_VERSION_REQUIRED").ToArray());
            }

            GoogleDriveFile remote = await driveClient.GetFileAsync(fileId, report: report, cancellationToken: cancellationToken).ConfigureAwait(false);
            if (remote.Capabilities != null && !remote.Capabilities.CanEdit) {
                throw new InvalidOperationException($"Drive file '{fileId}' cannot be edited by the current principal.");
            }

            if (replace.ConflictMode == GoogleSheetsReplaceConflictMode.RequireMatchingDriveVersion
                && remote.Version != replace.ExpectedDriveVersion) {
                string expected = replace.ExpectedDriveVersion!.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                string actual = remote.Version?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "<unknown>";
                report.Add(
                    TranslationSeverity.Error,
                    "ReplaceConflict",
                    $"Drive version changed from {expected} to {actual}; the spreadsheet was not mutated.",
                    code: "SHEETS.REPLACE.VERSION_CONFLICT",
                    action: TranslationAction.Fail,
                    targetId: fileId);
                throw new GoogleWorkspaceConflictException(
                    $"Google spreadsheet '{fileId}' changed after it was read.",
                    fileId,
                    expected,
                    actual,
                    report);
            }

            report.Add(
                replace.ConflictMode == GoogleSheetsReplaceConflictMode.Overwrite ? TranslationSeverity.Warning : TranslationSeverity.Info,
                "ReplaceConflict",
                replace.ConflictMode == GoogleSheetsReplaceConflictMode.Overwrite
                    ? "Destructive replacement was explicitly configured to overwrite the current remote version."
                    : "The current Drive version matches the caller's expected version.",
                code: replace.ConflictMode == GoogleSheetsReplaceConflictMode.Overwrite
                    ? "SHEETS.REPLACE.OVERWRITE"
                    : "SHEETS.REPLACE.VERSION_MATCH",
                action: TranslationAction.Preserve,
                targetId: fileId);
        }

        private static string? BuildSpreadsheetWebViewLink(string? spreadsheetId) {
            return string.IsNullOrWhiteSpace(spreadsheetId)
                ? null
                : $"https://docs.google.com/spreadsheets/d/{spreadsheetId}/edit";
        }

        private static async Task SendValuesAsync(
            GoogleWorkspaceHttpTransport transport,
            string accessToken,
            string spreadsheetId,
            GoogleSheetsBatch batch,
            IReadOnlyDictionary<string, int> sheetIds,
            GoogleSheetsExecutionOptions execution,
            CancellationToken cancellationToken) {
            if (!execution.UseValuesBatchUpdate || string.IsNullOrWhiteSpace(spreadsheetId)) return;
            GoogleSheetsApiBatchUpdateValuesPayload allValues = GoogleSheetsApiPayloadBuilder.BuildValuesBatchUpdatePayload(batch, sheetIds, spreadsheetId);
            int total = allValues.Data.Count;
            if (total == 0) return;

            int completed = 0;
            while (completed < total) {
                var payload = new GoogleSheetsApiBatchUpdateValuesPayload();
                payload.Data.AddRange(allValues.Data.Skip(completed).Take(execution.MaxValueRangesPerRequest));
                await transport.SendJsonAsync<GoogleSheetsApiBatchUpdateValuesPayload, object>(
                    accessToken,
                    HttpMethod.Post,
                    $"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheetId}/values:batchUpdate",
                    payload,
                    GoogleWorkspaceRequestSafety.Idempotent,
                    "Google Sheets API",
                    batch.Report,
                    GoogleSheetsJsonSerializerContext.Default.GoogleSheetsApiBatchUpdateValuesPayload,
                    GoogleSheetsJsonSerializerContext.Default.Object,
                    cancellationToken).ConfigureAwait(false);
                completed += payload.Data.Count;
                execution.Progress?.Report(new GoogleSheetsExportProgress("values", completed, total));
            }

            batch.Report.Add(
                TranslationSeverity.Info,
                "ValueBatching",
                $"Wrote {total} sparse value ranges through spreadsheets.values.batchUpdate.",
                code: "SHEETS.VALUES.BATCH_UPDATE",
                action: TranslationAction.Preserve,
                count: total,
                targetId: spreadsheetId);
        }

        private static async Task SendStructuralBatchesAsync(
            GoogleWorkspaceHttpTransport transport,
            string accessToken,
            string spreadsheetId,
            GoogleSheetsApiBatchUpdatePayload allRequests,
            GoogleSheetsExecutionOptions execution,
            TranslationReport report,
            CancellationToken cancellationToken) {
            int total = allRequests.Requests.Count;
            int completed = 0;
            while (completed < total) {
                var payload = new GoogleSheetsApiBatchUpdatePayload();
                payload.Requests.AddRange(allRequests.Requests.Skip(completed).Take(execution.MaxStructuralRequestsPerBatch));
                await transport.SendJsonAsync<GoogleSheetsApiBatchUpdatePayload, object>(
                    accessToken,
                    HttpMethod.Post,
                    $"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheetId}:batchUpdate",
                    payload,
                    GoogleWorkspaceRequestSafety.NonIdempotent,
                    "Google Sheets API",
                    report,
                    GoogleSheetsJsonSerializerContext.Default.GoogleSheetsApiBatchUpdatePayload,
                    GoogleSheetsJsonSerializerContext.Default.Object,
                    cancellationToken).ConfigureAwait(false);
                completed += payload.Requests.Count;
                execution.Progress?.Report(new GoogleSheetsExportProgress("structure", completed, total));
            }
        }

        private static void ValidateExecutionOptions(GoogleSheetsExecutionOptions execution) {
            if (execution == null) throw new ArgumentNullException(nameof(execution));
            if (execution.MaxValueRangesPerRequest <= 0) {
                throw new ArgumentOutOfRangeException(nameof(execution.MaxValueRangesPerRequest));
            }
            if (execution.MaxStructuralRequestsPerBatch <= 0) {
                throw new ArgumentOutOfRangeException(nameof(execution.MaxStructuralRequestsPerBatch));
            }
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
        public int? SheetId { get; set; }

        [System.Text.Json.Serialization.JsonPropertyName("title")]
        public string? Title { get; set; }
    }

}

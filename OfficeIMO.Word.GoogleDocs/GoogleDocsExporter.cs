using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using System.Text.Json;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Default Word to Google Docs exporter implementation.
    /// </summary>
    public sealed partial class GoogleDocsExporter : IGoogleDocsExporter {
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions {
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
            PropertyNamingPolicy = null,
            WriteIndented = false,
        };

        public GoogleDocsTranslationPlan BuildPlan(WordDocument document, GoogleDocsSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return GoogleDocsPlanBuilder.Build(document, options ?? new GoogleDocsSaveOptions());
        }

        public GoogleDocsBatch BuildBatch(WordDocument document, GoogleDocsSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            return GoogleDocsBatchCompiler.Build(document, options ?? new GoogleDocsSaveOptions());
        }

        public async Task<GoogleDocumentReference> ExportAsync(
            WordDocument document,
            GoogleWorkspaceSession session,
            GoogleDocsSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (session == null) throw new ArgumentNullException(nameof(session));

            var effectiveOptions = options ?? new GoogleDocsSaveOptions();
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
                accessToken = await session.AcquireAccessTokenAsync(GoogleWorkspaceScopeCatalog.DocsAuthoring, cancellationToken).ConfigureAwait(false);
            } catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested) {
                throw GoogleWorkspaceFailureDiagnostics.CreateRequestTimeoutFailure(
                    "Google Docs export token acquisition",
                    session.Options,
                    batch.Report,
                    ex);
            } catch (OperationCanceledException ex) when (cancellationToken.IsCancellationRequested) {
                throw GoogleWorkspaceFailureDiagnostics.CreateCanceledFailure(
                    "Google Docs export",
                    session.Options,
                    batch.Report,
                    ex);
            } catch (Exception ex) when (!(ex is OperationCanceledException)) {
                throw GoogleWorkspaceFailureDiagnostics.CreateTokenAcquisitionFailure(
                    "Google Docs export",
                    GoogleWorkspaceScopeCatalog.DocsAuthoring,
                    session,
                    batch.Report,
                    ex);
            }

            using (var transport = new GoogleWorkspaceHttpTransport(session.Options)) {
            using (var driveClient = new GoogleDriveClient(session)) {
            try {
                if (!string.IsNullOrWhiteSpace(effectiveLocation.ExistingFileId)) {
                    var existingDocument = await transport.SendJsonAsync<GoogleDocsApiDocumentResponse>(
                        accessToken.AccessToken,
                        HttpMethod.Get,
                        $"https://docs.googleapis.com/v1/documents/{effectiveLocation.ExistingFileId}?includeTabsContent=true",
                        null,
                        GoogleWorkspaceRequestSafety.Safe,
                        "Google Docs API",
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);

                    ConfigureExistingDocumentWrite(batch, existingDocument, effectiveOptions);

                    var resetPayload = GoogleDocsApiPayloadBuilder.BuildResetDocumentPayload(existingDocument, effectiveOptions.Tabs);
                    if (resetPayload.Requests.Count > 0) {
                        await SendBatchUpdateAsync(transport, accessToken.AccessToken, effectiveLocation.ExistingFileId!, batch, resetPayload, cancellationToken).ConfigureAwait(false);
                    }

                    await ApplyDocumentContentAsync(
                        transport,
                        accessToken.AccessToken,
                        effectiveLocation.ExistingFileId!,
                        batch,
                        effectiveOptions,
                        driveClient,
                        cancellationToken).ConfigureAwait(false);

                    await ApplyCommentsAsync(document, driveClient, effectiveLocation.ExistingFileId!, effectiveOptions, batch.Report, cancellationToken).ConfigureAwait(false);

                    var updatedDriveMetadata = await ApplyDrivePlacementAsync(
                        driveClient,
                        effectiveLocation.ExistingFileId!,
                        effectiveLocation,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);

                    batch.Report.Add(
                        TranslationSeverity.Info,
                        "ExistingDocument",
                        "Existing Google Docs replacement currently clears the body content before replaying the OfficeIMO batch.");

                    return new GoogleDocumentReference {
                        DocumentId = effectiveLocation.ExistingFileId,
                        FileId = effectiveLocation.ExistingFileId,
                        Name = existingDocument.Title ?? batch.Title,
                        MimeType = "application/vnd.google-apps.document",
                        WebViewLink = updatedDriveMetadata?.WebViewLink ?? BuildDocumentWebViewLink(effectiveLocation.ExistingFileId),
                        Location = effectiveLocation,
                        RevisionId = batch.WriteControlState?.RevisionId ?? existingDocument.RevisionId,
                        Report = batch.Report,
                    };
                }

                var createResponse = await transport.SendJsonAsync<GoogleDocsApiCreateDocumentResponse>(
                    accessToken.AccessToken,
                    HttpMethod.Post,
                    "https://docs.googleapis.com/v1/documents",
                    GoogleDocsApiPayloadBuilder.BuildCreateDocumentPayload(batch),
                    GoogleWorkspaceRequestSafety.NonIdempotent,
                    "Google Docs API",
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);

                if (string.IsNullOrWhiteSpace(createResponse.DocumentId)) {
                    throw new InvalidOperationException("Google Docs create response did not return a documentId.");
                }

                var documentId = createResponse.DocumentId!;
                ConfigureCreatedDocumentWrite(batch, createResponse, effectiveOptions);

                await ApplyDocumentContentAsync(
                    transport,
                    accessToken.AccessToken,
                    documentId,
                    batch,
                    effectiveOptions,
                    driveClient,
                    cancellationToken).ConfigureAwait(false);

                await ApplyCommentsAsync(document, driveClient, documentId, effectiveOptions, batch.Report, cancellationToken).ConfigureAwait(false);

                var createdDriveMetadata = await ApplyDrivePlacementAsync(
                    driveClient,
                    documentId,
                    effectiveLocation,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);

                return new GoogleDocumentReference {
                    DocumentId = documentId,
                    FileId = documentId,
                    Name = createResponse.Title ?? batch.Title,
                    MimeType = "application/vnd.google-apps.document",
                    WebViewLink = createdDriveMetadata?.WebViewLink ?? BuildDocumentWebViewLink(documentId),
                    Location = effectiveLocation,
                    RevisionId = batch.WriteControlState?.RevisionId ?? createResponse.RevisionId,
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
                    "Google Docs export",
                    session.Options,
                    batch.Report,
                    ex);
            } catch (OperationCanceledException ex) when (cancellationToken.IsCancellationRequested) {
                throw GoogleWorkspaceFailureDiagnostics.CreateCanceledFailure(
                    "Google Docs export",
                    session.Options,
                    batch.Report,
                    ex);
            } catch (Exception ex) when (!(ex is OperationCanceledException)) {
                throw GoogleWorkspaceFailureDiagnostics.CreateApiFailure(
                    "Google Docs export",
                    session.Options,
                    batch.Report,
                    ex);
            }
            }
            }
        }

        private static void ConfigureExistingDocumentWrite(
            GoogleDocsBatch batch,
            GoogleDocsApiDocumentResponse document,
            GoogleDocsSaveOptions options) {
            if (options.Tabs.Strategy == GoogleDocsTabStrategy.SelectedTab && string.IsNullOrWhiteSpace(options.Tabs.TabId)) {
                throw new ArgumentException("Tabs.TabId is required when SelectedTab is used.", nameof(options));
            }
            GoogleDocsApiTabResponse? target = GoogleDocsApiPayloadBuilder.SelectTabs(document, options.Tabs).FirstOrDefault();
            batch.TargetTabId = target?.Properties.TabId;

            string? expected = options.Replace.ExpectedRevisionId;
            if (options.Replace.ConflictMode != GoogleDocsRevisionConflictMode.OverwriteLatest && string.IsNullOrWhiteSpace(expected)) {
                batch.Report.Add(TranslationSeverity.Error, "ReplaceConflict", "Replacing a Google document requires the revision observed by a prior read/import.",
                    code: "DOCS.REPLACE.EXPECTED_REVISION_REQUIRED", action: TranslationAction.Fail, targetId: document.DocumentId);
                throw new GoogleWorkspacePreflightException(
                    "Google Docs replacement requires Replace.ExpectedRevisionId unless OverwriteLatest is explicitly selected.",
                    batch.Report,
                    batch.Report.Notices.Where(notice => notice.Code == "DOCS.REPLACE.EXPECTED_REVISION_REQUIRED").ToArray());
            }
            if (!string.IsNullOrWhiteSpace(expected) && !string.Equals(expected, document.RevisionId, StringComparison.Ordinal)) {
                batch.Report.Add(TranslationSeverity.Error, "ReplaceConflict", "The Google document revision changed after it was read.",
                    code: "DOCS.REPLACE.REVISION_CONFLICT", action: TranslationAction.Fail, targetId: document.DocumentId);
                throw new GoogleWorkspaceConflictException(
                    $"Google document '{document.DocumentId}' changed after it was read.",
                    document.DocumentId ?? "document",
                    expected,
                    document.RevisionId,
                    batch.Report);
            }
            batch.WriteControlState = new GoogleDocsWriteControlState(options.Replace.ConflictMode, expected ?? document.RevisionId);
        }

        private static void ConfigureCreatedDocumentWrite(
            GoogleDocsBatch batch,
            GoogleDocsApiCreateDocumentResponse document,
            GoogleDocsSaveOptions options) {
            batch.TargetTabId = options.Tabs.Strategy == GoogleDocsTabStrategy.SelectedTab
                ? options.Tabs.TabId
                : GoogleDocsApiPayloadBuilder.FlattenTabs(document.Tabs).FirstOrDefault()?.Properties.TabId;
            batch.WriteControlState = new GoogleDocsWriteControlState(GoogleDocsRevisionConflictMode.RequireRevision, document.RevisionId);
        }

    }
}

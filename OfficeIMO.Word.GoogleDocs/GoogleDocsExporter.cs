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
                        $"https://docs.googleapis.com/v1/documents/{effectiveLocation.ExistingFileId}",
                        null,
                        GoogleWorkspaceRequestSafety.Safe,
                        "Google Docs API",
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);

                    var resetPayload = GoogleDocsApiPayloadBuilder.BuildResetDocumentPayload(existingDocument);
                    if (resetPayload.Requests.Count > 0) {
                        await transport.SendJsonAsync<object>(
                            accessToken.AccessToken,
                            HttpMethod.Post,
                            $"https://docs.googleapis.com/v1/documents/{effectiveLocation.ExistingFileId}:batchUpdate",
                            resetPayload,
                            GoogleWorkspaceRequestSafety.NonIdempotent,
                            "Google Docs API",
                            batch.Report,
                            cancellationToken).ConfigureAwait(false);
                    }

                    await ApplyDocumentContentAsync(
                        transport,
                        accessToken.AccessToken,
                        effectiveLocation.ExistingFileId!,
                        batch,
                        effectiveOptions,
                        driveClient,
                        cancellationToken).ConfigureAwait(false);

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

                await ApplyDocumentContentAsync(
                    transport,
                    accessToken.AccessToken,
                    documentId,
                    batch,
                    effectiveOptions,
                    driveClient,
                    cancellationToken).ConfigureAwait(false);

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
                    Report = batch.Report,
                };
            } catch (GoogleWorkspaceExportException) {
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

    }
}

using OfficeIMO.GoogleWorkspace;
using System.Net.Http.Headers;
using System.IO;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// Default Word to Google Docs exporter implementation.
    /// </summary>
    public sealed class GoogleDocsExporter : IGoogleDocsExporter {
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

            var retryOptions = GoogleWorkspaceRetryOptions.FromSessionOptions(session.Options);

            bool disposeClient = session.Options.HttpClient == null;
            var client = session.Options.HttpClient ?? new HttpClient();
            try {
                client.Timeout = session.Options.RequestTimeout;

                if (!string.IsNullOrWhiteSpace(effectiveLocation.ExistingFileId)) {
                    var existingDocument = await SendAsync<GoogleDocsApiDocumentResponse>(
                        client,
                        accessToken.AccessToken,
                        HttpMethod.Get,
                        $"https://docs.googleapis.com/v1/documents/{effectiveLocation.ExistingFileId}",
                        null,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);

                    var resetPayload = GoogleDocsApiPayloadBuilder.BuildResetDocumentPayload(existingDocument);
                    if (resetPayload.Requests.Count > 0) {
                        await SendAsync<object>(
                            client,
                            accessToken.AccessToken,
                            HttpMethod.Post,
                            $"https://docs.googleapis.com/v1/documents/{effectiveLocation.ExistingFileId}:batchUpdate",
                            resetPayload,
                            retryOptions,
                            batch.Report,
                            cancellationToken).ConfigureAwait(false);
                    }

                    await ApplyDocumentContentAsync(
                        client,
                        accessToken.AccessToken,
                        effectiveLocation.ExistingFileId!,
                        batch,
                        retryOptions,
                        cancellationToken).ConfigureAwait(false);

                    var updatedDriveMetadata = await ApplyDrivePlacementAsync(
                        client,
                        accessToken.AccessToken,
                        effectiveLocation.ExistingFileId!,
                        effectiveLocation,
                        retryOptions,
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

                var createResponse = await SendAsync<GoogleDocsApiCreateDocumentResponse>(
                    client,
                    accessToken.AccessToken,
                    HttpMethod.Post,
                    "https://docs.googleapis.com/v1/documents",
                    GoogleDocsApiPayloadBuilder.BuildCreateDocumentPayload(batch),
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);

                if (string.IsNullOrWhiteSpace(createResponse.DocumentId)) {
                    throw new InvalidOperationException("Google Docs create response did not return a documentId.");
                }

                var documentId = createResponse.DocumentId!;

                await ApplyDocumentContentAsync(
                    client,
                    accessToken.AccessToken,
                    documentId,
                    batch,
                    retryOptions,
                    cancellationToken).ConfigureAwait(false);

                var createdDriveMetadata = await ApplyDrivePlacementAsync(
                    client,
                    accessToken.AccessToken,
                    documentId,
                    effectiveLocation,
                    retryOptions,
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
            } finally {
                if (disposeClient) {
                    client.Dispose();
                }
            }
        }

        private static async Task ApplyDocumentContentAsync(
            HttpClient client,
            string accessToken,
            string documentId,
            GoogleDocsBatch batch,
            GoogleWorkspaceRetryOptions retryOptions,
            CancellationToken cancellationToken) {
            var imageUris = await UploadInlineImagesAsync(
                client,
                accessToken,
                batch,
                retryOptions,
                cancellationToken).ConfigureAwait(false);

            var preparedInitialBatch = GoogleDocsApiPayloadBuilder.BuildPreparedInitialBatchUpdate(batch, imageUris);
            GoogleDocsApiBatchUpdateResponse? initialResponse = null;
            var initialPayload = preparedInitialBatch.Payload;
            if (initialPayload.Requests.Count > 0) {
                initialResponse = await SendAsync<GoogleDocsApiBatchUpdateResponse>(
                    client,
                    accessToken,
                    HttpMethod.Post,
                    $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                    initialPayload,
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);
            }

            if (preparedInitialBatch.Footnotes.Count > 0 && initialResponse != null) {
                await ApplyFootnotesAsync(
                    client,
                    accessToken,
                    documentId,
                    batch,
                    preparedInitialBatch.Footnotes,
                    initialResponse,
                    imageUris,
                    retryOptions,
                    cancellationToken).ConfigureAwait(false);
            }

            bool needsDocumentState = batch.Requests.OfType<GoogleDocsInsertTableRequest>().Any()
                || batch.Segments.Any(segment =>
                    string.Equals(segment.Variant, "default", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(segment.Variant, "first", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(segment.Variant, "even", StringComparison.OrdinalIgnoreCase));

            if (!needsDocumentState) {
                return;
            }

            var documentState = await SendAsync<GoogleDocsApiDocumentResponse>(
                client,
                accessToken,
                HttpMethod.Get,
                $"https://docs.googleapis.com/v1/documents/{documentId}",
                null,
                retryOptions,
                batch.Report,
                cancellationToken).ConfigureAwait(false);

            await ApplyHeaderFooterSegmentsAsync(
                client,
                accessToken,
                documentId,
                batch,
                imageUris,
                documentState,
                retryOptions,
                cancellationToken).ConfigureAwait(false);

            if (batch.Requests.OfType<GoogleDocsInsertTableRequest>().Any()) {
                var preparedTableBatch = GoogleDocsApiPayloadBuilder.BuildPreparedTableContentBatchUpdate(batch, documentState, imageUris);
                GoogleDocsApiBatchUpdateResponse? tableContentResponse = null;
                var tablePayload = preparedTableBatch.Payload;
                if (tablePayload.Requests.Count > 0) {
                    tableContentResponse = await SendAsync<GoogleDocsApiBatchUpdateResponse>(
                        client,
                        accessToken,
                        HttpMethod.Post,
                        $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                        tablePayload,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);
                }

                if (preparedTableBatch.Footnotes.Count > 0 && tableContentResponse != null) {
                    await ApplyFootnotesAsync(
                        client,
                        accessToken,
                        documentId,
                        batch,
                        preparedTableBatch.Footnotes,
                        tableContentResponse,
                        imageUris,
                        retryOptions,
                        cancellationToken).ConfigureAwait(false);
                }

                var mergePayload = GoogleDocsApiPayloadBuilder.BuildTableMergeBatchUpdatePayload(batch, documentState);
                if (mergePayload.Requests.Count > 0) {
                    await SendAsync<object>(
                        client,
                        accessToken,
                        HttpMethod.Post,
                        $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                        mergePayload,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);
                }

                var tableStylePayload = GoogleDocsApiPayloadBuilder.BuildTableStyleBatchUpdatePayload(batch, documentState);
                if (tableStylePayload.Requests.Count > 0) {
                    await SendAsync<object>(
                        client,
                        accessToken,
                        HttpMethod.Post,
                        $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                        tableStylePayload,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);
                }
            }
        }

        private static async Task<IReadOnlyDictionary<GoogleDocsInlineImage, string>> UploadInlineImagesAsync(
            HttpClient client,
            string accessToken,
            GoogleDocsBatch batch,
            GoogleWorkspaceRetryOptions retryOptions,
            CancellationToken cancellationToken) {
            var imageUris = new Dictionary<GoogleDocsInlineImage, string>();
            foreach (var image in EnumerateInlineImages(batch)) {
                if (!TryResolveImageUploadPayload(image, out var uploadName, out var mimeType, out var bytes, out var diagnosticMessage)) {
                    batch.Report.Add(
                        TranslationSeverity.Warning,
                        "InlineImages",
                        diagnosticMessage);
                    continue;
                }

                var fileId = await UploadDriveFileAsync(
                    client,
                    accessToken,
                    uploadName,
                    mimeType,
                    bytes,
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);
                await CreatePublicReadPermissionAsync(client, accessToken, fileId, retryOptions, batch.Report, cancellationToken).ConfigureAwait(false);
                imageUris[image] = BuildDrivePublicImageUri(fileId);
            }

            return imageUris;
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

        private static IEnumerable<GoogleDocsInlineImage> EnumerateInlineImages(GoogleDocsBatch batch) {
            var seen = new HashSet<GoogleDocsInlineImage>();
            foreach (var paragraphRequest in batch.Requests.OfType<GoogleDocsInsertParagraphRequest>()) {
                foreach (var image in EnumerateParagraphImages(paragraphRequest.Paragraph)) {
                    if (seen.Add(image)) {
                        yield return image;
                    }
                }
            }

            foreach (var tableRequest in batch.Requests.OfType<GoogleDocsInsertTableRequest>()) {
                foreach (var image in EnumerateTableImages(tableRequest.Table)) {
                    if (seen.Add(image)) {
                        yield return image;
                    }
                }
            }

            foreach (var segment in batch.Segments) {
                foreach (var request in segment.Requests) {
                    switch (request) {
                        case GoogleDocsInsertParagraphRequest paragraphRequest:
                            foreach (var image in EnumerateParagraphImages(paragraphRequest.Paragraph)) {
                                if (seen.Add(image)) {
                                    yield return image;
                                }
                            }
                            break;
                        case GoogleDocsInsertTableRequest tableRequest:
                            foreach (var image in EnumerateTableImages(tableRequest.Table)) {
                                if (seen.Add(image)) {
                                    yield return image;
                                }
                            }
                            break;
                    }
                }
            }
        }

        private static IEnumerable<GoogleDocsInlineImage> EnumerateParagraphImages(GoogleDocsParagraph paragraph) {
            foreach (var run in paragraph.Runs) {
                if (run.InlineImage != null) {
                    yield return run.InlineImage;
                }

                if (run.Footnote == null) {
                    continue;
                }

                foreach (var footnoteParagraph in run.Footnote.Paragraphs) {
                    foreach (var image in EnumerateParagraphImages(footnoteParagraph)) {
                        yield return image;
                    }
                }
            }
        }

        private static IEnumerable<GoogleDocsInlineImage> EnumerateTableImages(GoogleDocsTable table) {
            foreach (var row in table.Rows) {
                foreach (var cell in row.Cells) {
                    foreach (var paragraph in cell.Paragraphs) {
                        foreach (var image in EnumerateParagraphImages(paragraph)) {
                            yield return image;
                        }
                    }
                }
            }
        }

        private static async Task ApplyHeaderFooterSegmentsAsync(
            HttpClient client,
            string accessToken,
            string documentId,
            GoogleDocsBatch batch,
            IReadOnlyDictionary<GoogleDocsInlineImage, string> imageUris,
            GoogleDocsApiDocumentResponse documentState,
            GoogleWorkspaceRetryOptions retryOptions,
            CancellationToken cancellationToken) {
            var executableSegments = batch.Segments
                .Where(segment =>
                    string.Equals(segment.Variant, "default", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(segment.Variant, "first", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(segment.Variant, "even", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (executableSegments.Count == 0) {
                return;
            }

            var sectionBreakIndexes = EnumerateSectionBreakIndexes(documentState).ToList();
            foreach (var segment in executableSegments) {
                string? sectionBreakLocation = null;
                if (segment.SectionIndex > 0) {
                    if (sectionBreakIndexes.Count < segment.SectionIndex) {
                        batch.Report.Add(
                            TranslationSeverity.Warning,
                            "HeadersAndFooters",
                            $"Could not resolve the Google Docs section break location for section {segment.SectionIndex + 1}, so its {segment.Variant} {segment.Kind} was skipped.");
                        continue;
                    }

                    sectionBreakLocation = sectionBreakIndexes[segment.SectionIndex - 1].ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                string? segmentId;
                if (string.Equals(segment.Kind, "header", StringComparison.OrdinalIgnoreCase)) {
                    segmentId = await CreateHeaderAsync(client, accessToken, documentId, sectionBreakLocation, segment.Variant, retryOptions, batch.Report, cancellationToken).ConfigureAwait(false);
                } else {
                    segmentId = await CreateFooterAsync(client, accessToken, documentId, sectionBreakLocation, segment.Variant, retryOptions, batch.Report, cancellationToken).ConfigureAwait(false);
                }

                if (string.IsNullOrWhiteSpace(segmentId)) {
                    batch.Report.Add(
                        TranslationSeverity.Warning,
                        "HeadersAndFooters",
                        $"Google Docs did not return a segment id for section {segment.SectionIndex + 1} {segment.Variant} {segment.Kind}, so its content was skipped.");
                    continue;
                }

                var segmentPayload = GoogleDocsApiPayloadBuilder.BuildSegmentBatchUpdatePayload(segment, batch.Report, segmentId!, imageUris);
                if (segmentPayload.Requests.Count > 0) {
                    await SendAsync<object>(
                        client,
                        accessToken,
                        HttpMethod.Post,
                        $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                        segmentPayload,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);
                }

                if (!segment.Requests.OfType<GoogleDocsInsertTableRequest>().Any()) {
                    continue;
                }

                var segmentDocumentState = await SendAsync<GoogleDocsApiDocumentResponse>(
                    client,
                    accessToken,
                    HttpMethod.Get,
                    $"https://docs.googleapis.com/v1/documents/{documentId}",
                    null,
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);

                var segmentTablePayload = GoogleDocsApiPayloadBuilder.BuildSegmentTableContentBatchUpdatePayload(
                    segment,
                    segmentDocumentState,
                    batch.Report,
                    segmentId!,
                    imageUris);
                if (segmentTablePayload.Requests.Count == 0) {
                    continue;
                }

                await SendAsync<object>(
                    client,
                    accessToken,
                    HttpMethod.Post,
                    $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                    segmentTablePayload,
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);

                var segmentMergePayload = GoogleDocsApiPayloadBuilder.BuildSegmentTableMergeBatchUpdatePayload(
                    segment,
                    segmentDocumentState,
                    batch.Report,
                    segmentId!);
                if (segmentMergePayload.Requests.Count > 0) {
                    await SendAsync<object>(
                        client,
                        accessToken,
                        HttpMethod.Post,
                        $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                        segmentMergePayload,
                        retryOptions,
                        batch.Report,
                        cancellationToken).ConfigureAwait(false);
                }

                var segmentTableStylePayload = GoogleDocsApiPayloadBuilder.BuildSegmentTableStyleBatchUpdatePayload(
                    segment,
                    segmentDocumentState,
                    batch.Report,
                    segmentId!);
                if (segmentTableStylePayload.Requests.Count == 0) {
                    continue;
                }

                await SendAsync<object>(
                    client,
                    accessToken,
                    HttpMethod.Post,
                    $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                    segmentTableStylePayload,
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);
            }
        }

        private static async Task ApplyFootnotesAsync(
            HttpClient client,
            string accessToken,
            string documentId,
            GoogleDocsBatch batch,
            IReadOnlyList<GoogleDocsFootnote> footnotes,
            GoogleDocsApiBatchUpdateResponse initialResponse,
            IReadOnlyDictionary<GoogleDocsInlineImage, string> imageUris,
            GoogleWorkspaceRetryOptions retryOptions,
            CancellationToken cancellationToken) {
            var footnoteReplies = initialResponse.Replies
                .Where(reply => reply.CreateFootnote?.FootnoteId != null)
                .Select(reply => reply.CreateFootnote!.FootnoteId!)
                .ToList();

            if (footnoteReplies.Count != footnotes.Count) {
                batch.Report.Add(
                    TranslationSeverity.Warning,
                    "Footnotes",
                    $"Expected {footnotes.Count} Google Docs footnote replies after creation, but the API returned {footnoteReplies.Count}. Footnote content replay may be incomplete.");
            }

            for (int index = 0; index < Math.Min(footnotes.Count, footnoteReplies.Count); index++) {
                var footnotePayload = GoogleDocsApiPayloadBuilder.BuildFootnoteBatchUpdatePayload(
                    footnotes[index],
                    batch.Report,
                    footnoteReplies[index],
                    imageUris);
                if (footnotePayload.Requests.Count == 0) {
                    continue;
                }

                await SendAsync<object>(
                    client,
                    accessToken,
                    HttpMethod.Post,
                    $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                    footnotePayload,
                    retryOptions,
                    batch.Report,
                    cancellationToken).ConfigureAwait(false);
            }
        }

        private static IEnumerable<int> EnumerateSectionBreakIndexes(GoogleDocsApiDocumentResponse documentState) {
            var content = documentState.Body?.Content;
            if (content == null) {
                yield break;
            }

            foreach (var element in content) {
                if (element.SectionBreak != null && element.StartIndex.HasValue) {
                    yield return element.StartIndex.Value;
                }
            }
        }

        private static async Task<string?> CreateHeaderAsync(
            HttpClient client,
            string accessToken,
            string documentId,
            string? sectionBreakLocation,
            string variant,
            GoogleWorkspaceRetryOptions retryOptions,
            TranslationReport report,
            CancellationToken cancellationToken) {
            var payload = new GoogleDocsApiBatchUpdatePayload();
            payload.Requests.Add(new GoogleDocsApiRequestPayload {
                CreateHeader = new GoogleDocsApiCreateHeaderRequestPayload {
                    Type = ResolveHeaderFooterType(variant),
                    SectionBreakLocation = string.IsNullOrWhiteSpace(sectionBreakLocation)
                        ? null
                        : new GoogleDocsApiLocationPayload { Index = int.Parse(sectionBreakLocation, System.Globalization.CultureInfo.InvariantCulture) }
                }
            });

            var response = await SendAsync<GoogleDocsApiBatchUpdateResponse>(
                client,
                accessToken,
                HttpMethod.Post,
                $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                payload,
                retryOptions,
                report,
                cancellationToken).ConfigureAwait(false);

            return response.Replies.FirstOrDefault()?.CreateHeader?.HeaderId;
        }

        private static async Task<string?> CreateFooterAsync(
            HttpClient client,
            string accessToken,
            string documentId,
            string? sectionBreakLocation,
            string variant,
            GoogleWorkspaceRetryOptions retryOptions,
            TranslationReport report,
            CancellationToken cancellationToken) {
            var payload = new GoogleDocsApiBatchUpdatePayload();
            payload.Requests.Add(new GoogleDocsApiRequestPayload {
                CreateFooter = new GoogleDocsApiCreateFooterRequestPayload {
                    Type = ResolveHeaderFooterType(variant),
                    SectionBreakLocation = string.IsNullOrWhiteSpace(sectionBreakLocation)
                        ? null
                        : new GoogleDocsApiLocationPayload { Index = int.Parse(sectionBreakLocation, System.Globalization.CultureInfo.InvariantCulture) }
                }
            });

            var response = await SendAsync<GoogleDocsApiBatchUpdateResponse>(
                client,
                accessToken,
                HttpMethod.Post,
                $"https://docs.googleapis.com/v1/documents/{documentId}:batchUpdate",
                payload,
                retryOptions,
                report,
                cancellationToken).ConfigureAwait(false);

            return response.Replies.FirstOrDefault()?.CreateFooter?.FooterId;
        }

        private static string ResolveHeaderFooterType(string variant) {
            if (string.Equals(variant, "first", StringComparison.OrdinalIgnoreCase)) {
                return "FIRST_PAGE";
            }

            if (string.Equals(variant, "even", StringComparison.OrdinalIgnoreCase)) {
                return "EVEN_PAGE";
            }

            return "DEFAULT";
        }

        private static bool TryResolveImageUploadPayload(
            GoogleDocsInlineImage image,
            out string fileName,
            out string mimeType,
            out byte[] bytes,
            out string diagnosticMessage) {
            fileName = image.FileName ?? string.Empty;
            mimeType = image.ContentType ?? string.Empty;
            bytes = Array.Empty<byte>();
            diagnosticMessage = string.Empty;

            if (image.Bytes != null && image.Bytes.Length > 0) {
                bytes = image.Bytes;
                if (string.IsNullOrWhiteSpace(fileName)) {
                    fileName = string.IsNullOrWhiteSpace(image.FilePath) ? "officeimo-inline-image" : Path.GetFileName(image.FilePath);
                }

                if (string.IsNullOrWhiteSpace(mimeType) && !TryGetImageMimeType(fileName, out mimeType)) {
                    diagnosticMessage = "An inline image was embedded in the Word document, but its content type could not be inferred for the current Google Docs upload slice.";
                    return false;
                }

                return true;
            }

            if (!string.IsNullOrWhiteSpace(image.FilePath) && File.Exists(image.FilePath)) {
                var existingFilePath = image.FilePath!;
                fileName = string.IsNullOrWhiteSpace(fileName) ? Path.GetFileName(existingFilePath) : fileName;
                if (string.IsNullOrWhiteSpace(mimeType) && !TryGetImageMimeType(existingFilePath, out mimeType)) {
                    diagnosticMessage = $"Inline image '{existingFilePath}' uses an unsupported extension for the current Google Docs image upload slice, so export kept the readable placeholder.";
                    return false;
                }

                bytes = File.ReadAllBytes(existingFilePath);
                return true;
            }

            if (!string.IsNullOrWhiteSpace(image.FilePath)) {
                diagnosticMessage = $"Inline image file '{image.FilePath}' was not found, so Google Docs export kept the readable placeholder instead of a native image.";
                return false;
            }

            diagnosticMessage = "A Word inline image did not expose embedded bytes or a local file path, so Google Docs export kept the readable placeholder instead of a native image.";
            return false;
        }

        private static bool TryGetImageMimeType(string filePath, out string mimeType) {
            switch (Path.GetExtension(filePath).ToLowerInvariant()) {
                case ".png":
                    mimeType = "image/png";
                    return true;
                case ".jpg":
                case ".jpeg":
                    mimeType = "image/jpeg";
                    return true;
                case ".gif":
                    mimeType = "image/gif";
                    return true;
                case ".bmp":
                    mimeType = "image/bmp";
                    return true;
                default:
                    mimeType = string.Empty;
                    return false;
            }
        }

        private static async Task<string> UploadDriveFileAsync(
            HttpClient client,
            string accessToken,
            string fileName,
            string mimeType,
            byte[] fileBytes,
            GoogleWorkspaceRetryOptions retryOptions,
            TranslationReport report,
            CancellationToken cancellationToken) {
            var response = await SendAsync<GoogleDriveFileMetadataResponse>(
                client,
                accessToken,
                HttpMethod.Post,
                "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id",
                () => {
                    var boundary = "officeimo-" + Guid.NewGuid().ToString("N");
                    var metadataJson = JsonSerializer.Serialize(new {
                        name = fileName,
                        mimeType,
                    }, JsonOptions);

                    var content = new MultipartContent("related", boundary);
                    var metadataContent = new StringContent(metadataJson, Encoding.UTF8, "application/json");
                    content.Add(metadataContent);

                    var fileContent = new ByteArrayContent(fileBytes);
                    fileContent.Headers.ContentType = new MediaTypeHeaderValue(mimeType);
                    content.Add(fileContent);
                    return content;
                },
                retryOptions,
                report,
                cancellationToken).ConfigureAwait(false);

            if (string.IsNullOrWhiteSpace(response.Id)) {
                throw new InvalidOperationException("Drive image upload did not return a file id.");
            }

            return response.Id!;
        }

        private static Task<object> CreatePublicReadPermissionAsync(
            HttpClient client,
            string accessToken,
            string fileId,
            GoogleWorkspaceRetryOptions retryOptions,
            TranslationReport report,
            CancellationToken cancellationToken) {
            return SendAsync<object>(
                client,
                accessToken,
                HttpMethod.Post,
                $"https://www.googleapis.com/drive/v3/files/{fileId}/permissions?supportsAllDrives=true",
                new {
                    role = "reader",
                    type = "anyone",
                },
                retryOptions,
                report,
                cancellationToken);
        }

        private static string BuildDrivePublicImageUri(string fileId) {
            return "https://drive.google.com/uc?export=download&id=" + Uri.EscapeDataString(fileId);
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

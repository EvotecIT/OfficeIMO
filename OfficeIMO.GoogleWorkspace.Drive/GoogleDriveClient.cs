using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.GoogleWorkspace.Drive {
    /// <summary>OAuth scopes, shared-drive behavior, and download limits for <see cref="GoogleDriveClient"/>.</summary>
    public sealed class GoogleDriveClientOptions {
        /// <summary>Default maximum buffered or file download size: 256 MiB.</summary>
        public const long DefaultMaxDownloadBytes = 256L * 1024L * 1024L;
        /// <summary>Gets or sets OAuth scopes used by read operations.</summary>
        public IReadOnlyList<string> ReadScopes { get; set; } = new[] { GoogleWorkspaceScopeCatalog.DriveReadonly };
        /// <summary>Gets or sets OAuth scopes used by mutation operations.</summary>
        public IReadOnlyList<string> WriteScopes { get; set; } = new[] { GoogleWorkspaceScopeCatalog.DriveFile };
        /// <summary>Gets or sets whether requests include shared-drive support. The default is <see langword="true"/>.</summary>
        public bool SupportsAllDrives { get; set; } = true;
        /// <summary>Gets or sets the maximum response body accepted by download operations.</summary>
        public long MaxDownloadBytes { get; set; } = DefaultMaxDownloadBytes;

        /// <summary>
        /// Creates an option set for an authoring workflow that only reads files created or opened by the app.
        /// </summary>
        public static GoogleDriveClientOptions ForFileAuthoring() {
            return new GoogleDriveClientOptions {
                ReadScopes = new[] { GoogleWorkspaceScopeCatalog.DriveFile },
                WriteScopes = new[] { GoogleWorkspaceScopeCatalog.DriveFile },
            };
        }
    }

    /// <summary>Filtering, paging, projection, and corpus options for listing Drive files.</summary>
    public sealed class GoogleDriveListOptions {
        /// <summary>Gets or sets the Drive query expression.</summary>
        public string? Query { get; set; }
        /// <summary>Gets or sets a shared-drive identifier to restrict the corpus.</summary>
        public string? DriveId { get; set; }
        /// <summary>Gets or sets the continuation token from a previous page.</summary>
        public string? PageToken { get; set; }
        /// <summary>Gets or sets the requested page size, clamped to 1 through 1000.</summary>
        public int PageSize { get; set; } = 100;
        /// <summary>Gets or sets a comma-separated list of spaces to search. The default is <c>drive</c>.</summary>
        public string Spaces { get; set; } = "drive";
        /// <summary>Gets or sets the Drive API ordering expression.</summary>
        public string? OrderBy { get; set; }
        /// <summary>Gets or sets the partial-response fields expression.</summary>
        public string? Fields { get; set; }
        /// <summary>Gets or sets whether results may include items from shared drives.</summary>
        public bool IncludeItemsFromAllDrives { get; set; } = true;
    }

    /// <summary>Dependency-light client for Google Drive metadata, content, collaboration, and change operations.</summary>
    public sealed partial class GoogleDriveClient : IDisposable {
        /// <summary>Default Drive fields requested for file metadata.</summary>
        public const string DefaultFileFields = "id,name,mimeType,driveId,parents,webViewLink,webContentLink,modifiedTime,createdTime,version,size,trashed,capabilities(canDownload,canEdit,canMoveItemWithinDrive,canMoveItemOutOfDrive,canDelete,canShare,canComment)";

        private readonly GoogleWorkspaceSession _session;
        private readonly GoogleWorkspaceHttpTransport _transport;
        private readonly GoogleDriveClientOptions _options;
        private bool _disposed;

        /// <summary>Creates a Drive client bound to a Workspace session and a snapshot of client options.</summary>
        public GoogleDriveClient(GoogleWorkspaceSession session, GoogleDriveClientOptions? options = null) {
            _session = session ?? throw new ArgumentNullException(nameof(session));
            GoogleDriveClientOptions configured = options ?? new GoogleDriveClientOptions();
            if (configured.MaxDownloadBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Maximum Google Drive download bytes must be positive.");
            _options = new GoogleDriveClientOptions {
                ReadScopes = SnapshotScopes(configured.ReadScopes, nameof(GoogleDriveClientOptions.ReadScopes)),
                WriteScopes = SnapshotScopes(configured.WriteScopes, nameof(GoogleDriveClientOptions.WriteScopes)),
                SupportsAllDrives = configured.SupportsAllDrives,
                MaxDownloadBytes = configured.MaxDownloadBytes,
            };
            _transport = new GoogleWorkspaceHttpTransport(session);
        }

        /// <summary>Returns the least-privilege default scope for a logical Drive operation.</summary>
        public static IReadOnlyList<string> GetRequiredScopes(GoogleDriveOperation operation) {
            switch (operation) {
                case GoogleDriveOperation.ReadMetadata:
                case GoogleDriveOperation.Download:
                case GoogleDriveOperation.Export:
                case GoogleDriveOperation.ReadComments:
                case GoogleDriveOperation.ReadRevisions:
                case GoogleDriveOperation.ReadChanges:
                    return new[] { GoogleWorkspaceScopeCatalog.DriveReadonly };
                case GoogleDriveOperation.CreateOrUpdate:
                case GoogleDriveOperation.ManagePermissions:
                case GoogleDriveOperation.ManageComments:
                case GoogleDriveOperation.Delete:
                    return new[] { GoogleWorkspaceScopeCatalog.DriveFile };
                default:
                    throw new ArgumentOutOfRangeException(nameof(operation));
            }
        }

        /// <summary>Gets the MIME formats Drive can import and export for the current account.</summary>
        public async Task<GoogleDriveAboutFormats> GetFormatsAsync(
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(_options.ReadScopes, report, "Google Drive format discovery", cancellationToken).ConfigureAwait(false);
            return await _transport.SendJsonAsync<GoogleDriveAboutFormats>(
                token,
                HttpMethod.Get,
                "https://www.googleapis.com/drive/v3/about?fields=importFormats,exportFormats",
                null,
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveAboutFormats,
                cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Gets metadata for one Drive file using the default or caller-selected field projection.</summary>
        public async Task<GoogleDriveFile> GetFileAsync(
            string fileId,
            string? fields = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateId(fileId, nameof(fileId));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(_options.ReadScopes, report, "Google Drive file metadata", cancellationToken).ConfigureAwait(false);
            string uri = $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}?fields={Escape(fields ?? DefaultFileFields)}&supportsAllDrives={Bool(_options.SupportsAllDrives)}";
            return await _transport.SendJsonAsync<GoogleDriveFile>(
                token,
                HttpMethod.Get,
                uri,
                null,
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveFile,
                cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Lists one page of Drive files using the requested query and corpus options.</summary>
        public async Task<GoogleDriveFileList> ListFilesAsync(
            GoogleDriveListOptions? options = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            options ??= new GoogleDriveListOptions();
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(_options.ReadScopes, report, "Google Drive file listing", cancellationToken).ConfigureAwait(false);

            var query = new List<string> {
                "pageSize=" + Math.Max(1, Math.Min(1000, options.PageSize)).ToString(System.Globalization.CultureInfo.InvariantCulture),
                "spaces=" + Escape(options.Spaces),
                "supportsAllDrives=" + Bool(_options.SupportsAllDrives),
                "includeItemsFromAllDrives=" + Bool(options.IncludeItemsFromAllDrives),
                "fields=" + Escape(options.Fields ?? $"nextPageToken,incompleteSearch,files({DefaultFileFields})"),
            };
            AddQuery(query, "q", options.Query);
            AddQuery(query, "pageToken", options.PageToken);
            AddQuery(query, "orderBy", options.OrderBy);

            if (!string.IsNullOrWhiteSpace(options.DriveId)) {
                query.Add("corpora=drive");
                query.Add("driveId=" + Escape(options.DriveId!));
            }

            return await _transport.SendJsonAsync<GoogleDriveFileList>(
                token,
                HttpMethod.Get,
                "https://www.googleapis.com/drive/v3/files?" + string.Join("&", query),
                null,
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveFileList,
                cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Creates a Drive folder, optionally under a parent folder.</summary>
        public async Task<GoogleDriveFile> CreateFolderAsync(
            string name,
            string? parentId = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Folder name is required.", nameof(name));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(_options.WriteScopes, report, "Google Drive folder creation", cancellationToken).ConfigureAwait(false);
            var payload = GoogleDriveJson.ToNode(new GoogleDriveFilePayload {
                Name = name,
                MimeType = GoogleDriveMimeTypes.Folder,
                Parents = string.IsNullOrWhiteSpace(parentId) ? null : new[] { parentId! },
            }, GoogleDriveJsonSerializerContext.Default.GoogleDriveFilePayload);
            string uri = $"https://www.googleapis.com/drive/v3/files?supportsAllDrives={Bool(_options.SupportsAllDrives)}&fields={Escape(DefaultFileFields)}";
            return await _transport.SendJsonAsync<GoogleDriveFile>(
                token,
                HttpMethod.Post,
                uri,
                payload,
                GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveFile,
                cancellationToken,
                mutationKind: GoogleWorkspaceMutationKind.Create,
                requiredScopes: _options.WriteScopes).ConfigureAwait(false);
        }

        /// <summary>Copies a Drive file with an optional new name and parent.</summary>
        public async Task<GoogleDriveFile> CopyFileAsync(
            string fileId,
            string? name = null,
            string? parentId = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateId(fileId, nameof(fileId));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(_options.WriteScopes, report, "Google Drive file copy", cancellationToken).ConfigureAwait(false);
            var payload = GoogleDriveJson.ToNode(new GoogleDriveFilePayload {
                Name = name,
                Parents = string.IsNullOrWhiteSpace(parentId) ? null : new[] { parentId! },
            }, GoogleDriveJsonSerializerContext.Default.GoogleDriveFilePayload);
            string uri = $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}/copy?supportsAllDrives={Bool(_options.SupportsAllDrives)}&fields={Escape(DefaultFileFields)}";
            return await _transport.SendJsonAsync<GoogleDriveFile>(
                token,
                HttpMethod.Post,
                uri,
                payload,
                GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveFile,
                cancellationToken,
                mutationKind: GoogleWorkspaceMutationKind.Create,
                requiredScopes: _options.WriteScopes).ConfigureAwait(false);
        }

        /// <summary>Moves a file to one folder, removing its existing parents when necessary.</summary>
        public async Task<GoogleDriveFile> MoveFileAsync(
            string fileId,
            string folderId,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateId(fileId, nameof(fileId));
            ValidateId(folderId, nameof(folderId));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(_options.WriteScopes, report, "Google Drive file move", cancellationToken).ConfigureAwait(false);
            var current = await GetFileWithTokenAsync(token, fileId, DefaultFileFields, report, cancellationToken).ConfigureAwait(false);
            if (current.Parents.Count == 1 && string.Equals(current.Parents[0], folderId, StringComparison.Ordinal)) {
                return current;
            }

            var query = new List<string> {
                "supportsAllDrives=" + Bool(_options.SupportsAllDrives),
                "addParents=" + Escape(folderId),
                "fields=" + Escape(DefaultFileFields),
            };
            if (current.Parents.Count > 0) {
                query.Add("removeParents=" + Escape(string.Join(",", current.Parents)));
            }

            return await _transport.SendJsonAsync<GoogleDriveFile>(
                token,
                new HttpMethod("PATCH"),
                $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}?{string.Join("&", query)}",
                new System.Text.Json.Nodes.JsonObject(),
                GoogleWorkspaceRequestSafety.Idempotent,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveFile,
                cancellationToken,
                requiredScopes: _options.WriteScopes).ConfigureAwait(false);
        }

        /// <summary>Permanently deletes a Drive file using guarded mutation policy.</summary>
        public async Task DeleteFileAsync(
            string fileId,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateId(fileId, nameof(fileId));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(_options.WriteScopes, report, "Google Drive file deletion", cancellationToken).ConfigureAwait(false);
            await DeleteFileWithTokenAsync(token, fileId, report, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Gets metadata and caller capabilities for a shared drive.</summary>
        public async Task<GoogleSharedDrive> GetSharedDriveAsync(
            string driveId,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateId(driveId, nameof(driveId));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(_options.ReadScopes, report, "Google shared drive metadata", cancellationToken).ConfigureAwait(false);
            return await _transport.SendJsonAsync<GoogleSharedDrive>(
                token,
                HttpMethod.Get,
                $"https://www.googleapis.com/drive/v3/drives/{Escape(driveId)}?fields=id,name,hidden,createdTime,capabilities",
                null,
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleSharedDrive,
                cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Gets a folder and verifies its type and, when supplied, containing shared drive.</summary>
        public async Task<GoogleDriveFile> ResolveFolderAsync(
            string folderId,
            string? expectedDriveId = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            var folder = await GetFileAsync(folderId, DefaultFileFields, report, cancellationToken).ConfigureAwait(false);
            if (!string.Equals(folder.MimeType, GoogleDriveMimeTypes.Folder, StringComparison.Ordinal)) {
                throw new InvalidOperationException($"Google Drive item '{folderId}' is not a folder.");
            }

            if (!string.IsNullOrWhiteSpace(expectedDriveId)
                && !string.Equals(folder.DriveId, expectedDriveId, StringComparison.Ordinal)) {
                throw new InvalidOperationException($"Google Drive folder '{folderId}' belongs to drive '{folder.DriveId ?? "My Drive"}', not expected drive '{expectedDriveId}'.");
            }

            return folder;
        }

        /// <summary>Releases transport resources owned by the client.</summary>
        public void Dispose() {
            if (_disposed) return;
            _transport.Dispose();
            _disposed = true;
        }

        internal async Task<string> AcquireTokenAsync(
            IReadOnlyList<string> scopes,
            TranslationReport report,
            string operationName,
            CancellationToken cancellationToken) {
            try {
                var token = await _session.AcquireAccessTokenAsync(scopes, cancellationToken).ConfigureAwait(false);
                return token.AccessToken;
            } catch (Exception exception) when (!(exception is OperationCanceledException)) {
                throw GoogleWorkspaceFailureDiagnostics.CreateTokenAcquisitionFailure(operationName, scopes, _session, report, exception);
            }
        }

        internal Task<GoogleDriveFile> GetFileWithTokenAsync(
            string token,
            string fileId,
            string fields,
            TranslationReport report,
            CancellationToken cancellationToken) {
            string uri = $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}?fields={Escape(fields)}&supportsAllDrives={Bool(_options.SupportsAllDrives)}";
            return _transport.SendJsonAsync(
                token,
                HttpMethod.Get,
                uri,
                null,
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveFile,
                cancellationToken);
        }

        internal Task DeleteFileWithTokenAsync(
            string token,
            string fileId,
            TranslationReport report,
            CancellationToken cancellationToken) {
            return _transport.SendJsonAsync<object>(
                token,
                HttpMethod.Delete,
                $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}?supportsAllDrives={Bool(_options.SupportsAllDrives)}",
                null,
                GoogleWorkspaceRequestSafety.Idempotent,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.Object,
                cancellationToken,
                revisionPrecondition: GoogleWorkspaceRevisionPrecondition.Unavailable,
                requiredScopes: _options.WriteScopes);
        }

        internal GoogleWorkspaceHttpTransport Transport => _transport;
        internal GoogleDriveClientOptions Options => _options;

        internal static string Escape(string value) => Uri.EscapeDataString(value);
        internal static string Bool(bool value) => value ? "true" : "false";

        private static void ValidateId(string value, string parameterName) {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A Google Drive identifier is required.", parameterName);
        }

        private static IReadOnlyList<string> SnapshotScopes(
            IReadOnlyList<string>? scopes,
            string optionName) {
            if (scopes == null || scopes.Count == 0 || scopes.Any(string.IsNullOrWhiteSpace)) {
                throw new ArgumentException("Google Drive scopes must contain non-empty OAuth scope values.", optionName);
            }
            return Array.AsReadOnly(scopes.Distinct(StringComparer.Ordinal).ToArray());
        }

        private static void AddQuery(ICollection<string> query, string name, string? value) {
            if (!string.IsNullOrWhiteSpace(value)) query.Add(Escape(name) + "=" + Escape(value!));
        }
    }

    /// <summary>Logical Drive operation used to select a least-privilege OAuth scope.</summary>
    public enum GoogleDriveOperation {
        /// <summary>Read file, folder, or shared-drive metadata.</summary>
        ReadMetadata = 0,
        /// <summary>Download binary file content.</summary>
        Download = 1,
        /// <summary>Export a Google-native file to an external format.</summary>
        Export = 2,
        /// <summary>Create, copy, move, or upload content.</summary>
        CreateOrUpdate = 3,
        /// <summary>Create or delete file permissions.</summary>
        ManagePermissions = 4,
        /// <summary>Read file comments and replies.</summary>
        ReadComments = 5,
        /// <summary>Create, update, resolve, or delete comments and replies.</summary>
        ManageComments = 6,
        /// <summary>Read file revision history.</summary>
        ReadRevisions = 7,
        /// <summary>Read the account or shared-drive change feed.</summary>
        ReadChanges = 8,
        /// <summary>Permanently delete a Drive file.</summary>
        Delete = 9,
    }
}

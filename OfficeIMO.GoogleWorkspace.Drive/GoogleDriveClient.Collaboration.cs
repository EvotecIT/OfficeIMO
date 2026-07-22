using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.GoogleWorkspace.Drive {
    public sealed class GoogleDrivePermissionCreateOptions {
        public string Type { get; set; } = "user";
        public string Role { get; set; } = "reader";
        public string? EmailAddress { get; set; }
        public string? Domain { get; set; }
        public bool? AllowFileDiscovery { get; set; }
        public bool SendNotificationEmail { get; set; } = true;
        public string? EmailMessage { get; set; }
    }

    public sealed class GoogleDriveChangeListOptions {
        public string? DriveId { get; set; }
        public int PageSize { get; set; } = 100;
        public bool IncludeRemoved { get; set; } = true;
        public bool IncludeCorpusRemovals { get; set; } = true;
        /// <summary>
        /// Whether a user change feed should include items from shared drives. Disable this when shared-drive
        /// feeds are consumed separately to prevent the same shared-drive change from being observed twice.
        /// </summary>
        public bool IncludeItemsFromAllDrives { get; set; } = true;
        public string? Fields { get; set; }
    }

    public sealed partial class GoogleDriveClient {
        /// <summary>Lists one page of permissions for a file or shared drive.</summary>
        /// <param name="fileId">The file or shared-drive identifier.</param>
        /// <param name="pageToken">The continuation token returned by the previous page.</param>
        /// <param name="pageSize">The requested number of permissions. Google Drive caps this value at 100.</param>
        /// <param name="report">An optional translation report that receives API diagnostics.</param>
        /// <param name="cancellationToken">A token that cancels the request.</param>
        /// <returns>The requested permission page and its continuation token, when another page exists.</returns>
        public async Task<GoogleDrivePermissionList> ListPermissionsAsync(
            string fileId,
            string? pageToken = null,
            int? pageSize = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateResourceId(fileId, nameof(fileId));
            if (pageSize <= 0) throw new ArgumentOutOfRangeException(nameof(pageSize), "Page size must be positive when specified.");
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.ReadScopes, report, "Google Drive permission listing", cancellationToken).ConfigureAwait(false);
            var query = new List<string> {
                "supportsAllDrives=" + Bool(Options.SupportsAllDrives),
                "fields=nextPageToken,permissions(id,type,role,emailAddress,domain,displayName,allowFileDiscovery)",
            };
            if (!string.IsNullOrWhiteSpace(pageToken)) query.Add("pageToken=" + Escape(pageToken!));
            if (pageSize != null) query.Add("pageSize=" + pageSize.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
            return await Transport.SendJsonAsync<GoogleDrivePermissionList>(
                token,
                HttpMethod.Get,
                $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}/permissions?{string.Join("&", query)}",
                null,
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDrivePermissionList,
                cancellationToken).ConfigureAwait(false);
        }

        public async Task<GoogleDrivePermission> CreatePermissionAsync(
            string fileId,
            GoogleDrivePermissionCreateOptions options,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateResourceId(fileId, nameof(fileId));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (string.IsNullOrWhiteSpace(options.Type)) throw new ArgumentException("Permission type is required.", nameof(options));
            if (string.IsNullOrWhiteSpace(options.Role)) throw new ArgumentException("Permission role is required.", nameof(options));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.WriteScopes, report, "Google Drive permission creation", cancellationToken).ConfigureAwait(false);
            var payload = GoogleDriveJson.ToNode(new GoogleDrivePermissionPayload {
                Type = options.Type,
                Role = options.Role,
                EmailAddress = options.EmailAddress,
                Domain = options.Domain,
                AllowFileDiscovery = options.AllowFileDiscovery,
            }, GoogleDriveJsonSerializerContext.Default.GoogleDrivePermissionPayload);
            var query = new List<string> {
                "supportsAllDrives=" + Bool(Options.SupportsAllDrives),
                "sendNotificationEmail=" + Bool(options.SendNotificationEmail),
                "fields=id,type,role,emailAddress,domain,displayName,allowFileDiscovery",
            };
            if (!string.IsNullOrWhiteSpace(options.EmailMessage)) query.Add("emailMessage=" + Escape(options.EmailMessage!));
            return await Transport.SendJsonAsync<GoogleDrivePermission>(
                token,
                HttpMethod.Post,
                $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}/permissions?{string.Join("&", query)}",
                payload,
                GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDrivePermission,
                cancellationToken).ConfigureAwait(false);
        }

        public async Task DeletePermissionAsync(
            string fileId,
            string permissionId,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateResourceId(fileId, nameof(fileId));
            ValidateResourceId(permissionId, nameof(permissionId));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.WriteScopes, report, "Google Drive permission deletion", cancellationToken).ConfigureAwait(false);
            await Transport.SendJsonAsync<object>(
                token,
                HttpMethod.Delete,
                $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}/permissions/{Escape(permissionId)}?supportsAllDrives={Bool(Options.SupportsAllDrives)}",
                null,
                GoogleWorkspaceRequestSafety.Idempotent,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.Object,
                cancellationToken).ConfigureAwait(false);
        }

        public async Task<GoogleDriveCommentList> ListCommentsAsync(
            string fileId,
            string? pageToken = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateResourceId(fileId, nameof(fileId));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.ReadScopes, report, "Google Drive comment listing", cancellationToken).ConfigureAwait(false);
            string page = string.IsNullOrWhiteSpace(pageToken) ? string.Empty : "&pageToken=" + Escape(pageToken!);
            return await Transport.SendJsonAsync<GoogleDriveCommentList>(
                token,
                HttpMethod.Get,
                $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}/comments?includeDeleted=true&fields=nextPageToken,comments(id,content,anchor,resolved,deleted,createdTime,modifiedTime,replies(id,content,action,deleted,createdTime)){page}",
                null,
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveCommentList,
                cancellationToken).ConfigureAwait(false);
        }

        public async Task<GoogleDriveComment> CreateCommentAsync(
            string fileId,
            string content,
            string? anchor = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateResourceId(fileId, nameof(fileId));
            if (string.IsNullOrWhiteSpace(content)) throw new ArgumentException("Comment content is required.", nameof(content));
            report ??= new TranslationReport();
            if (!string.IsNullOrWhiteSpace(anchor)) {
                report.Add(
                    TranslationSeverity.Warning,
                    "Comments",
                    "Drive preserves the custom comment anchor, but Google Workspace editors display Drive API comments as unanchored.",
                    code: "DRIVE.COMMENT.EDITOR_ANCHOR_UNAVAILABLE",
                    action: TranslationAction.Preserve);
            }

            string token = await AcquireTokenAsync(Options.WriteScopes, report, "Google Drive comment creation", cancellationToken).ConfigureAwait(false);
            return await Transport.SendJsonAsync<GoogleDriveComment>(
                token,
                HttpMethod.Post,
                $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}/comments?fields=id,content,anchor,resolved,deleted,createdTime,modifiedTime,replies",
                GoogleDriveJson.ToNode(new GoogleDriveCommentPayload {
                    Content = content,
                    Anchor = anchor,
                }, GoogleDriveJsonSerializerContext.Default.GoogleDriveCommentPayload),
                GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveComment,
                cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Deletes a Drive comment thread from a file.</summary>
        public async Task DeleteCommentAsync(
            string fileId,
            string commentId,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateResourceId(fileId, nameof(fileId));
            ValidateResourceId(commentId, nameof(commentId));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.WriteScopes, report, "Google Drive comment deletion", cancellationToken).ConfigureAwait(false);
            await Transport.SendJsonAsync<object>(
                token,
                HttpMethod.Delete,
                $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}/comments/{Escape(commentId)}",
                null,
                GoogleWorkspaceRequestSafety.Idempotent,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.Object,
                cancellationToken).ConfigureAwait(false);
        }

        public async Task<GoogleDriveReply> CreateReplyAsync(
            string fileId,
            string commentId,
            string content,
            string? action = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateResourceId(fileId, nameof(fileId));
            ValidateResourceId(commentId, nameof(commentId));
            if (string.IsNullOrWhiteSpace(content) && string.IsNullOrWhiteSpace(action)) {
                throw new ArgumentException("Reply content or an action is required.", nameof(content));
            }

            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.WriteScopes, report, "Google Drive comment reply creation", cancellationToken).ConfigureAwait(false);
            return await Transport.SendJsonAsync<GoogleDriveReply>(
                token,
                HttpMethod.Post,
                $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}/comments/{Escape(commentId)}/replies?fields=id,content,action,deleted,createdTime",
                GoogleDriveJson.ToNode(new GoogleDriveReplyPayload {
                    Content = content,
                    Action = action,
                }, GoogleDriveJsonSerializerContext.Default.GoogleDriveReplyPayload),
                GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveReply,
                cancellationToken).ConfigureAwait(false);
        }

        /// <summary>Deletes a reply from a Drive comment thread.</summary>
        public async Task DeleteReplyAsync(
            string fileId,
            string commentId,
            string replyId,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateResourceId(fileId, nameof(fileId));
            ValidateResourceId(commentId, nameof(commentId));
            ValidateResourceId(replyId, nameof(replyId));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.WriteScopes, report, "Google Drive comment reply deletion", cancellationToken).ConfigureAwait(false);
            await Transport.SendJsonAsync<object>(
                token,
                HttpMethod.Delete,
                $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}/comments/{Escape(commentId)}/replies/{Escape(replyId)}",
                null,
                GoogleWorkspaceRequestSafety.Idempotent,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.Object,
                cancellationToken).ConfigureAwait(false);
        }

        public async Task<GoogleDriveRevisionList> ListRevisionsAsync(
            string fileId,
            string? pageToken = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            ValidateResourceId(fileId, nameof(fileId));
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.ReadScopes, report, "Google Drive revision listing", cancellationToken).ConfigureAwait(false);
            string page = string.IsNullOrWhiteSpace(pageToken) ? string.Empty : "&pageToken=" + Escape(pageToken!);
            return await Transport.SendJsonAsync<GoogleDriveRevisionList>(
                token,
                HttpMethod.Get,
                $"https://www.googleapis.com/drive/v3/files/{Escape(fileId)}/revisions?fields=nextPageToken,revisions(id,modifiedTime,keepForever,published,size){page}",
                null,
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveRevisionList,
                cancellationToken).ConfigureAwait(false);
        }

        public async Task<string> GetStartPageTokenAsync(
            string? driveId = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.ReadScopes, report, "Google Drive change-token acquisition", cancellationToken).ConfigureAwait(false);
            string drive = string.IsNullOrWhiteSpace(driveId) ? string.Empty : "&driveId=" + Escape(driveId!);
            var response = await Transport.SendJsonAsync<GoogleDriveStartPageToken>(
                token,
                HttpMethod.Get,
                $"https://www.googleapis.com/drive/v3/changes/startPageToken?supportsAllDrives={Bool(Options.SupportsAllDrives)}{drive}",
                null,
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveStartPageToken,
                cancellationToken).ConfigureAwait(false);
            return response.Value ?? throw new InvalidOperationException("Google Drive did not return a start page token.");
        }

        public async Task<GoogleDriveChangeList> ListChangesAsync(
            string pageToken,
            GoogleDriveChangeListOptions? options = null,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(pageToken)) throw new ArgumentException("A change page token is required.", nameof(pageToken));
            options ??= new GoogleDriveChangeListOptions();
            report ??= new TranslationReport();
            string token = await AcquireTokenAsync(Options.ReadScopes, report, "Google Drive change listing", cancellationToken).ConfigureAwait(false);
            var query = new List<string> {
                "pageToken=" + Escape(pageToken),
                "pageSize=" + Math.Max(1, Math.Min(1000, options.PageSize)).ToString(System.Globalization.CultureInfo.InvariantCulture),
                "includeRemoved=" + Bool(options.IncludeRemoved),
                "includeCorpusRemovals=" + Bool(options.IncludeCorpusRemovals),
                "includeItemsFromAllDrives=" + Bool(options.IncludeItemsFromAllDrives),
                "supportsAllDrives=" + Bool(Options.SupportsAllDrives),
                "fields=" + Escape(options.Fields ?? $"nextPageToken,newStartPageToken,changes(fileId,removed,changeType,driveId,file({DefaultFileFields}))"),
            };
            if (!string.IsNullOrWhiteSpace(options.DriveId)) query.Add("driveId=" + Escape(options.DriveId!));
            return await Transport.SendJsonAsync<GoogleDriveChangeList>(
                token,
                HttpMethod.Get,
                "https://www.googleapis.com/drive/v3/changes?" + string.Join("&", query),
                null,
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report,
                GoogleDriveJsonSerializerContext.Default.GoogleDriveChangeList,
                cancellationToken).ConfigureAwait(false);
        }

        private static void ValidateResourceId(string value, string parameterName) {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A Google Drive identifier is required.", parameterName);
        }
    }
}

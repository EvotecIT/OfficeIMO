# OfficeIMO.GoogleWorkspace.Drive

Typed, dependency-light Google Drive API support shared by the OfficeIMO Google Docs, Sheets, and Slides packages.

The package covers file and folder metadata, shared-drive validation, copy/move/delete, permissions, comments and replies, revisions, change tokens, import/export format discovery, downloads, multipart and resumable uploads, and temporary public-content leases with cleanup reporting. DELETE operations require a policy that explicitly accepts the named data loss; durable downloads verify and truncate only uncheckpointed crash tails, and resumable uploads revalidate their seekable source before reporting success.

Callers provide a `GoogleWorkspaceSession`; applications remain responsible for OAuth consent, credentials, and choosing scopes appropriate to the files they manage.
Durable transfers are stream/file based. Persist `GoogleDriveResumableUploadCheckpoint.Value` after every callback and protect it like a credential because it contains Google's upload-session URI. `GoogleDriveDownloadCheckpoint` is non-secret and binds a partial destination to the Drive file version, size, path identity, and committed-content hash. Both resume paths reconcile remote state before continuing and reject changed local or remote inputs.

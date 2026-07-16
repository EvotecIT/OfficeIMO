# OfficeIMO.GoogleWorkspace.Drive

Typed, dependency-light Google Drive API support shared by the OfficeIMO Google Docs, Sheets, and Slides packages.

The package covers file and folder metadata, shared-drive validation, copy/move/delete, permissions, comments and replies, revisions, change tokens, import/export format discovery, downloads, multipart and resumable uploads, and temporary public-content leases with cleanup reporting.

Callers provide a `GoogleWorkspaceSession`; applications remain responsible for OAuth consent, credentials, and choosing scopes appropriate to the files they manage.

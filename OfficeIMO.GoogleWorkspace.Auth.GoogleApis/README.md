# OfficeIMO.GoogleWorkspace.Auth.GoogleApis

Optional adapters for applications that already use `Google.Apis.Auth`. The package keeps the shared OfficeIMO Google Workspace kernel free of Google SDK dependencies, supports `GoogleCredential` and `UserCredential`, and provides an installed-application authorization entry point that always enables PKCE.

Interactive authorization requires an application-provided `IGoogleWorkspaceTokenStore`. OfficeIMO intentionally provides no plaintext refresh-token store by default.

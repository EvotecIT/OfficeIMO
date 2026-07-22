---
title: Authentication
description: Choose static, delegate, service-account, or Google APIs credential sources without leaking token policy into translators.
order: 60
---

`OfficeIMO.GoogleWorkspace` depends only on `IGoogleWorkspaceCredentialSource`. It includes static/delegate sources and a service-account JWT source with optional domain-wide delegation.

```csharp
var options = new GoogleWorkspaceSessionOptions {
    SubjectUser = "analyst@example.com",
    UseDomainWideDelegation = true
};
var credentials = GoogleServiceAccountCredentialSource.FromFile("service-account.json", options);
var session = new GoogleWorkspaceSession(credentials, options);
```

Install `OfficeIMO.GoogleWorkspace.Auth.GoogleApis` when an application already uses `Google.Apis.Auth`. `GoogleApisCredentialSource` adapts `GoogleCredential`, `UserCredential`, or `ITokenAccess`. `GoogleInstalledApplicationAuthorization` enables PKCE and requires an application-provided `IGoogleWorkspaceTokenStore`.

OfficeIMO deliberately provides no plaintext refresh-token store. The application must encrypt authorization state at rest, protect client secrets, choose the consent experience, and select the smallest scopes that satisfy its workflow. Never commit access tokens, service-account private keys, OAuth client secrets, or refresh tokens.

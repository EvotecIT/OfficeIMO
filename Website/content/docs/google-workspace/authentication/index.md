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

Install `OfficeIMO.GoogleWorkspace.Auth.GoogleApis` when an application already uses `Google.Apis.Auth`. `GoogleApisCredentialSource` adapts `GoogleCredential`, `UserCredential`, or `ITokenAccess`. For mutation-capable sessions, supply a `GoogleWorkspaceCredentialBindingResolver` that verifies the access token with Google and returns its provider-issued account and required grants. `GoogleInstalledApplicationAuthorization.AuthorizeAsync` enables PKCE and requires that resolver plus an application-provided `IGoogleWorkspaceTokenStore`.

`GoogleWorkspaceAccessToken.FromVerifiedCredential` keeps provider evidence separate from caller-entered policy. `GoogleWorkspaceSession` verifies that evidence against `ExpectedAccount` and the exact operation scopes, and mutation transport must be constructed from that session. An account label passed to a legacy or raw-token constructor remains informational and cannot make the token mutation-capable. This prevents a policy for one account from authorizing a token acquired for another account.

OfficeIMO deliberately provides no plaintext refresh-token store. The application must encrypt authorization state at rest, protect client secrets, choose the consent experience, and select the smallest scopes that satisfy its workflow. Never commit access tokens, service-account private keys, OAuth client secrets, or refresh tokens.

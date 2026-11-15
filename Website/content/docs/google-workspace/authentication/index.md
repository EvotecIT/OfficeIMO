---
title: "Authenticate to Google Workspace"
description: "Create a GoogleWorkspaceSession in .NET from a service account, a short-lived access token, or existing Google APIs credentials."
order: 60
packages:
  - OfficeIMO.GoogleWorkspace
  - OfficeIMO.GoogleWorkspace.Auth.GoogleApis
cmdlets:
  - New-OfficeGoogleWorkspaceSession
limits: "OfficeIMO provides no plaintext refresh-token store; the application must encrypt authorization state at rest."
search_terms:
  - google service account
  - domain-wide delegation
  - google oauth token
  - google workspace credentials
verified: 2026-11-14
---

Every translator works through a `GoogleWorkspaceSession`, which gets tokens from an `IGoogleWorkspaceCredentialSource`. `OfficeIMO.GoogleWorkspace` includes a service-account source with optional domain-wide delegation, a static token source, and a delegate source. `OfficeIMO.GoogleWorkspace.Auth.GoogleApis` adapts `Google.Apis.Auth` credentials for applications that already use them.

## Use a service account

The service-account source is the usual choice for services that create or change files. With domain-wide delegation, `SubjectUser` is the user the service account acts as.

### C#

```csharp
using OfficeIMO.GoogleWorkspace;

var options = new GoogleWorkspaceSessionOptions {
    ApplicationName = "Reporting service",
    SubjectUser = "analyst@example.com",
    UseDomainWideDelegation = true
};

using var credentials = GoogleServiceAccountCredentialSource.FromFile("service-account.json", options);
var session = new GoogleWorkspaceSession(credentials, options);
```

### PowerShell

```powershell
$session = New-OfficeGoogleWorkspaceSession -KeyPath .\Secrets\reporting.json -Subject analyst@example.com -Scope Documents, DriveFile
```

## Use a short-lived access token

A static token is enough for read-only samples and short jobs. An account name or scope supplied by the caller is not proof of what the token can do, so a static token cannot authorize changes. For changes from a delegate or the Google APIs adapter, return `GoogleWorkspaceAccessToken.FromVerifiedCredential` after checking the provider-issued token evidence.

### C#

```csharp
using OfficeIMO.GoogleWorkspace;

string token = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_ACCESS_TOKEN")
    ?? throw new InvalidOperationException("Set GOOGLE_WORKSPACE_ACCESS_TOKEN.");

var session = new GoogleWorkspaceSession(
    new StaticAccessTokenCredentialSource(token, DateTimeOffset.UtcNow.AddMinutes(50)),
    new GoogleWorkspaceSessionOptions { ApplicationName = "Read-only report" });
```

## Use Google APIs credentials

Install `OfficeIMO.GoogleWorkspace.Auth.GoogleApis` when the application already uses `Google.Apis.Auth`. `GoogleApisCredentialSource` adapts `GoogleCredential`, `UserCredential`, or `ITokenAccess`. For sessions that change files, supply a `GoogleWorkspaceCredentialBindingResolver` that verifies the token with Google and returns its account and grants. `GoogleInstalledApplicationAuthorization.AuthorizeAsync` runs the installed-application flow with PKCE and needs that resolver plus an `IGoogleWorkspaceTokenStore` your application provides.

`GoogleWorkspaceSession` checks the verified evidence against `ExpectedAccount` and the exact scopes of each operation, so a policy written for one account cannot authorize a token issued to another.

Keep access tokens, service-account keys, OAuth client secrets, and refresh tokens out of source control, and request the smallest scopes the workflow needs.

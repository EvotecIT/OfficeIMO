# OfficeIMO.GoogleWorkspace - shared Google Workspace primitives

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.GoogleWorkspace)](https://www.nuget.org/packages/OfficeIMO.GoogleWorkspace)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.GoogleWorkspace?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.GoogleWorkspace)

`OfficeIMO.GoogleWorkspace` contains the dependency-light credential, session, transport, retry, scope, diagnostics, Drive-location, and translation-report contracts shared by the OfficeIMO Google Docs, Sheets, Slides, Drive, and synchronization packages.

## Explicit mutation policy

Non-safe Google requests are blocked unless the session names the expected account, supplies an operation-policy provider, and records every outcome through a receipt sink. The policy captures scopes, target, expected revision decision, retry count and total elapsed-time deadline, rate-limit behavior, and the caller's data-loss decision. Transport-known destructive operations such as DELETE are refused unless the policy explicitly accepts the named loss. This applies in the dependency-light HTTP owner, so Docs, Sheets, Slides, Drive, and optional SDK adapters cannot bypass it accidentally.

Translation preflight defaults to `FailOnErrors`. Select `FailOnWarnings` when every lossy projection must be accepted by diagnostic code before mutation; a named accepted diagnostic remains explicit and reviewable.

Drive offers stream/file durable transfer APIs. `UploadResumableStreamAsync` and `UploadResumableFileAsync` persist the sensitive upload-session checkpoint after initiation and every confirmed chunk, query Google before resuming, reconcile ambiguous outcomes, and verify that the local source did not change. `DownloadToFileAsync` uses ranged reads and binds its checkpoint to file id, Drive version, size, destination identity, and the hash of committed bytes; after a crash it verifies the checkpointed prefix and discards only an uncheckpointed tail. Protect upload checkpoint values like credentials; their `ToString()` representation is redacted and policy/receipt targets use a stable SHA-256 session identifier.

## Install

```powershell
dotnet add package OfficeIMO.GoogleWorkspace
```

## Quick start

```csharp
using OfficeIMO.GoogleWorkspace;

var receipts = new List<GoogleWorkspaceOperationReceipt>();
var options = new GoogleWorkspaceSessionOptions {
    ApplicationName = "OfficeIMO Samples",
    ExpectedAccount = "author@example.com",
    DefaultDriveId = "shared-drive-id",
    DefaultFolderId = "reports-folder-id",
    MaxRetryCount = 5,
    RetryBaseDelay = TimeSpan.FromMilliseconds(250),
    RetryMaxDelay = TimeSpan.FromSeconds(10),
    MaxRetryElapsedTime = TimeSpan.FromMinutes(2),
    RequestTimeout = TimeSpan.FromSeconds(120),
    OperationReceiptSink = receipts.Add,
};
options.OperationPolicyProvider = context => {
    string expectedRevision = context.RevisionPreconditionKind switch {
        GoogleWorkspaceRevisionPreconditionKind.ResourceAbsentCreate => GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision,
        GoogleWorkspaceRevisionPreconditionKind.PayloadRevision => context.AdapterExpectedRevision!,
        GoogleWorkspaceRevisionPreconditionKind.ResumableSessionState => context.AdapterExpectedRevision!,
        GoogleWorkspaceRevisionPreconditionKind.Unavailable => GoogleWorkspaceOperationPolicy.ExplicitlyUnversionedRevision("API exposes no conditional revision"),
        _ => "\"observed-google-etag\"", // sent as If-Match
    };
    // Secure baseline: destructive or deliberately unversioned mutations stay blocked.
    // Applications should accept only specifically approved operations and targets here.
    bool acceptsLoss = false;
    return new GoogleWorkspaceOperationPolicy(
        options.ExpectedAccount!,
        context.RequiredScopes,
        context.Target,
        expectedRevision,
        context.MaxRetryCount,
        context.MaxRetryElapsedTime,
        context.RateLimitPolicy,
        acceptsLoss ? GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss : GoogleWorkspaceDataLossDecision.RejectPotentialLoss,
        acceptsLoss ? "application-approved operation and target" : null);
};

var session = new GoogleWorkspaceSession(
    new StaticAccessTokenCredentialSource("<google-access-token>"), options);
```

## What it provides

- `IGoogleWorkspaceCredentialSource` for application-owned OAuth or service-account token acquisition.
- `StaticAccessTokenCredentialSource`, `DelegateGoogleWorkspaceCredentialSource`, and `GoogleServiceAccountCredentialSource`.
- `GoogleWorkspaceSession` and `GoogleWorkspaceSessionOptions`.
- `GoogleDriveFileLocation` for folder, shared-drive, and existing-file targeting.
- `TranslationReport`, `TranslationNotice`, preflight, conflict/export exceptions, cancellation exceptions, and log-ready diagnostic entries.
- Safety-aware retries, normalized Google API failures, request timeouts, and diagnostic correlation through `GoogleWorkspaceSessionOptions.DiagnosticSink`.
- Minimum-scope catalogs for Docs, Sheets, Slides, and Drive operations.

## Service account shortcut

```csharp
var sessionOptions = new GoogleWorkspaceSessionOptions {
    SubjectUser = "analyst@example.com",
    UseDomainWideDelegation = true,
    DefaultFolderId = "reports-folder-id",
};

var credentialSource = GoogleServiceAccountCredentialSource.FromFile(
    "service-account.json",
    sessionOptions);

var session = new GoogleWorkspaceSession(credentialSource, sessionOptions);
```

For a Google endpoint that exposes no usable conditional revision precondition, call
`GoogleWorkspaceOperationPolicy.ExplicitlyUnversionedRevision(reason)` and pair it with
`AcceptSpecifiedLoss` plus a named accepted-loss description. The receipt then records that the mutation was
deliberately unguarded instead of presenting an observed version as an enforced precondition.
For Docs and Slides write-control payloads, `AdapterExpectedRevision` is the exact revision already embedded by
the adapter; return that same value so the transport can reject a mismatched policy before sending the request.

This shortcut is sufficient for read-only calls. Add the explicit mutation policy and receipt sink shown above before creating, updating, or deleting cloud resources.

## Boundaries

- This package owns shared Google Workspace plumbing.
- Word/Google Docs translation belongs in `OfficeIMO.Word.GoogleDocs`.
- Excel/Google Sheets translation belongs in `OfficeIMO.Excel.GoogleSheets`.
- PowerPoint/Google Slides translation belongs in `OfficeIMO.PowerPoint.GoogleSlides`.
- Drive resources belong in `OfficeIMO.GoogleWorkspace.Drive`; change-feed consumption and plan/apply belong in `OfficeIMO.GoogleWorkspace.Sync`.
- Applications own consent and credential policy. `OfficeIMO.GoogleWorkspace.Auth.GoogleApis` is an optional adapter when an application already uses the Google client SDK.

## Related packages

- [OfficeIMO.GoogleWorkspace.Drive](../OfficeIMO.GoogleWorkspace.Drive/README.md)
- [OfficeIMO.GoogleWorkspace.Auth.GoogleApis](../OfficeIMO.GoogleWorkspace.Auth.GoogleApis/README.md)
- [OfficeIMO.GoogleWorkspace.Sync](../OfficeIMO.GoogleWorkspace.Sync/README.md)
- [OfficeIMO.Word.GoogleDocs](../OfficeIMO.Word.GoogleDocs/README.md)
- [OfficeIMO.Excel.GoogleSheets](../OfficeIMO.Excel.GoogleSheets/README.md)
- [OfficeIMO.PowerPoint.GoogleSlides](../OfficeIMO.PowerPoint.GoogleSlides/README.md)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, plus `net472` on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** `System.Text.Json` plus platform HTTP and cryptography; no Google client SDK.
- **OfficeIMO:** Credential abstractions, sessions, retry, Drive placement, failures, and translation reports are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

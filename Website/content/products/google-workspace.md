---
title: "OfficeIMO Google Workspace"
description: "Translate Word, Excel, and PowerPoint content to and from Google Docs, Sheets, and Slides with explicit fidelity and conflict handling."
layout: product
product_color: "#4285f4"
install: "dotnet add package OfficeIMO.GoogleWorkspace"
nuget: "OfficeIMO.GoogleWorkspace"
docs_url: "/docs/google-workspace/"
api_url: "/api/google-workspace/"
---

## Office and Google Workspace without a second document model

OfficeIMO keeps Word, Excel, and PowerPoint as the source document models. Small companion libraries translate those models to Google Docs, Sheets, and Slides, while shared packages own authentication adapters, Drive resources, change feeds, retries, diagnostics, and plan/apply outcomes.

## Package map

| Package | Responsibility |
|---|---|
| `OfficeIMO.GoogleWorkspace` | Session, credential contract, scopes, transport, retries, diagnostics, fidelity preflight |
| `OfficeIMO.GoogleWorkspace.Drive` | Files, folders, shared drives, conversion, media, permissions, comments, revisions, change tokens |
| `OfficeIMO.GoogleWorkspace.Auth.GoogleApis` | Optional adapters for `Google.Apis.Auth`, PKCE, and application-owned encrypted token storage |
| `OfficeIMO.GoogleWorkspace.Sync` | Change-feed consumption, minimal checkpoints, dry-run/approval/apply outcomes |
| `OfficeIMO.Word.GoogleDocs` | Bidirectional Word and Google Docs translation |
| `OfficeIMO.Excel.GoogleSheets` | Bidirectional Excel and Google Sheets translation |
| `OfficeIMO.PowerPoint.GoogleSlides` | Bidirectional PowerPoint and Google Slides translation |

## The workflow

1. Build a translation plan without network access.
2. Review stable diagnostics and choose explicit fidelity fallbacks.
3. Create a new Google file, or import/read an existing file to obtain revision/version evidence.
4. Replace only with that evidence, unless the application explicitly chooses last-writer-wins.
5. Use Drive change feeds and the sync executor when the application needs repeatable synchronization.

Native imports favor editable Google semantics. Drive Office-format export is the broad-fidelity fallback. Neither path claims pixel-perfect parity where the platforms have different document models.

## Safety boundaries

- Non-idempotent mutations are not retried as if they were safe reads.
- Existing Docs, Sheets, and Slides replacements require explicit conflict policy.
- Lossy mappings are reported before mutation and can be blocked by fidelity preflight.
- Temporary image URLs are owned leases with cleanup on success and failure.
- OAuth client secrets, consent, refresh-token encryption, tenant policy, and checkpoint persistence remain application responsibilities.

See the [Google Workspace guides](/docs/google-workspace/) and the [generated support matrix](/docs/google-workspace/support/).

## Compatibility

The libraries target .NET Standard 2.0, .NET 8, .NET 10, and .NET Framework 4.7.2 on Windows. The dependency-light core does not require the Google client SDK; install the authentication adapter only when the application wants it.

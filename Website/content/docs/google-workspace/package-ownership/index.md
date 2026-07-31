---
title: Google Workspace package ownership
description: "Choose the OfficeIMO Google Workspace packages and use explicit fidelity, import, replacement, and synchronization contracts."
order: 90
aliases:
  - /docs/google-workspace/migration/
---

OfficeIMO separates transport, Drive operations, authentication, synchronization, and document-format mapping so applications depend only on the capabilities they use.

## Package map

| Package | Current owner |
| --- | --- |
| `OfficeIMO.GoogleWorkspace` | Sessions, scopes, safety-aware HTTP transport, retry/error normalization, diagnostics, fidelity policy, and shared Drive target references |
| `OfficeIMO.GoogleWorkspace.Drive` | Files, folders, shared drives, metadata, permissions, comments, revisions, change tokens, download/upload, and temporary content leases |
| `OfficeIMO.GoogleWorkspace.Auth.GoogleApis` | Optional Google SDK credential adapters, installed-application PKCE, and the application-owned token-store boundary |
| `OfficeIMO.GoogleWorkspace.Sync` | User and shared-drive change feeds, minimal checkpoints, dry-run plans, approvals, and apply outcomes |
| `OfficeIMO.Word.GoogleDocs` | Word/Google Docs planning, import, guarded replacement, comments, fidelity reports, and fallbacks |
| `OfficeIMO.Excel.GoogleSheets` | Excel/Google Sheets planning, import, version-safe replacement, formula/advanced-object mapping, and fidelity reports |
| `OfficeIMO.PowerPoint.GoogleSlides` | PowerPoint/Google Slides planning, import, template/guarded replacement, speaker notes, visual fallbacks, and fidelity reports |

## Fidelity and unsupported features

Use `UnsupportedFeatures`, `GoogleWorkspaceFidelityPolicy`, and the format support catalogs to decide whether a mapping is native, simplified, rasterized, preserved through a fallback, or blocked. A successful HTTP request is not a fidelity claim.

## Existing-file updates

Import or read the target before replacement. Docs and Slides use the observed API revision; Sheets uses the observed Drive version. Pass that evidence to the replacement operation. Choose an overwrite mode only when last-writer-wins is an application decision.

## Imports and image safety

Use native import when editable target semantics matter and Drive Office-format export for the documented broad fallback. Google Docs temporary image publication is explicit; Google Slides owns short-lived image leases internally. The adapters do not create permanent public staging objects.

## Synchronization boundary

The synchronization package owns change-feed cursors, plan/apply outcomes, and bounded execution. Applications own checkpoint persistence, schedules, approval UI, conflict decisions, and the mutation callback.

The current feature catalog is published in the [support matrix](/docs/google-workspace/support/).

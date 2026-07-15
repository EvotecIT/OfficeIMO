# OfficeIMO Google Workspace Libraries

## Current outcome

The Google-oriented code is now a package family with one reusable foundation, one Drive owner, optional authentication and synchronization packages, and thin format translators:

| Package | Owns |
|---|---|
| `OfficeIMO.GoogleWorkspace` | Credential/session contracts, scope catalog, safety-aware HTTP transport, retry/error normalization, diagnostics, fidelity policy, Drive target references |
| `OfficeIMO.GoogleWorkspace.Drive` | Files, folders, shared drives, metadata/capabilities, copy/move/delete, permissions, comments/replies, revisions, change tokens, conversion, download/upload, temporary content leases |
| `OfficeIMO.GoogleWorkspace.Auth.GoogleApis` | Optional `Google.Apis.Auth` credential adapters, installed-application PKCE, application-owned token-store boundary |
| `OfficeIMO.GoogleWorkspace.Sync` | User/per-drive change-feed consumption, minimal checkpoints, dry-run/approval/apply outcomes |
| `OfficeIMO.Word.GoogleDocs` | Word/Docs planning, create, tab-aware guarded replacement, native/Drive import, comments, fallbacks, diff/checkpoint |
| `OfficeIMO.Excel.GoogleSheets` | Excel/Sheets planning, create, version-safe replacement, formula/advanced-object mapping, native/Drive import, diff/checkpoint |
| `OfficeIMO.PowerPoint.GoogleSlides` | PowerPoint/Slides planning, create/template/guarded replacement, native/Drive import, speaker notes, visual fallback, diff/checkpoint |

This split is intentional. Document semantics remain in Word, Excel, and PowerPoint. Google transport and Drive behavior are not copied into the format packages. Synchronization does not become a second document database.

## End-to-end contract

Each format follows the same lifecycle:

1. Build a local translation plan.
2. Inspect stable diagnostics and enforce `GoogleWorkspaceFidelityPolicy` before mutation.
3. Create a new Google file, or import/read an existing file to obtain revision/version evidence.
4. Replace with explicit conflict policy.
5. Import natively for editable Google semantics or through Drive Office-format conversion for broad fallback.
6. Build a format diff when an application needs synchronization.
7. Use Drive change cursors and `OfficeIMO.GoogleWorkspace.Sync` for discovery, dry-run, approval, apply, cancellation, and partial outcomes.

Create/replace/import results return the Google file ID, URL, translation report, Drive version and modified time where available, and Docs/Slides revision where the API exposes it.

## Implemented safety decisions

- HTTP retries are classified by request safety; ambiguous non-idempotent mutations are not blindly replayed.
- Google API errors retain status, reason/category, request correlation, and translation diagnostics.
- `ApplicationName` and request correlation are sent through the shared transport.
- Existing Docs and Slides replacements require an observed revision by default. Sheets requires the observed Drive version.
- Docs image publication is explicit. Slides image publication uses an owned short-lived lease. Both clean up on failure paths.
- Lossy features carry stable diagnostic codes and executed actions. Strict preflight can stop before the first Google mutation.
- Shared-drive folder identity is validated rather than inferred from query flags.
- Change tracking advances each user/shared-drive cursor only after that source reaches a new start token.
- OAuth consent, client secrets, refresh-token encryption, and tenant policy remain application responsibilities.

## Format support

The authoritative format rows live in:

- `GoogleDocsFeatureSupportCatalog.Features`
- `GoogleSheetsFeatureSupportCatalog.Features`
- `GoogleSlidesFeatureSupportCatalog.Features`

The website matrix is generated from those classes:

```powershell
dotnet run --project OfficeIMO.Examples -f net10.0 -- `
  --google-support-matrix Website/content/docs/google-workspace/support/index.md
```

The generated page explains `Native`, `Partial`, `Flattened`, `Rasterized`, `DriveFallback`, and `Unsupported`. Documentation must not claim a capability that is absent from these catalogs and the corresponding translator tests.

## Validation lanes

Mocked contract tests cover payloads, retries, scope requests, conflict checks, cleanup, import projection, change pagination, dry-run/approval, cancellation, and partial failure on .NET 8, .NET 10, and .NET Framework 4.7.2. Package validation also builds .NET Standard 2.0 assets.

Disposable live tests exist for Docs, Sheets, Slides, and Drive change tracking. They are opt-in through:

- `OFFICEIMO_RUN_GOOGLE_WORKSPACE_LIVE=1`
- `GOOGLE_WORKSPACE_ACCESS_TOKEN`
- `GOOGLE_WORKSPACE_FOLDER_ID`
- optional `GOOGLE_WORKSPACE_DRIVE_ID`

The live lanes create resources only in the configured folder and delete them in `finally` paths.

## Completed implementation train

- [x] Replace misleading preview options with executed fidelity policies.
- [x] Extract shared transport, retry, scopes, diagnostics, and failure contracts.
- [x] Add the typed Drive owner, temporary leases, collaboration resources, conversion, media, and change tokens.
- [x] Add optional service-account/installed-application authentication paths without forcing Google SDK dependencies into core.
- [x] Complete Sheets create/replace/import, formulas, formats, advanced objects, chunking, identity, diff, and live contracts.
- [x] Complete Docs tabs, revision controls, full reset semantics, comments, import, fallbacks, diff, and live contracts.
- [x] Add the standalone Slides translator with native core, complex-slide fallback, notes, templates, import, diff, and live contracts.
- [x] Add user/shared-drive change consumption and generic plan/apply outcomes.
- [x] Add package READMEs, website product/guides, preview migration notes, runnable examples, and a generated support matrix.

## Deliberate boundaries

- Pixel-perfect equivalence between Microsoft Office and Google editors is not promised.
- Native Google import is not treated as broader than Drive conversion; each has a documented purpose.
- Google Sheets smart chips and embedded drawings are not inferred from arbitrary links or `IMAGE()` formulas.
- Full Google Docs representation of Word OLE, watermarks, equations, SmartArt, and floating layout is not invented.
- Full Slides masters, themes, transitions, animations, diagrams, equations, OLE, and media semantics are not invented; the default is a reported visual fallback for complex slides.
- Drive custom anchors are not advertised as native editor comment anchors.
- Fine-grained automatic merging is application policy, not a hidden translator side effect.

## Release checklist

Before publishing the family:

- [ ] Run all Google-focused tests on every supported test target.
- [ ] Pack every publishable Google package and inspect the NuGet assets/README.
- [ ] Run the disposable live lanes with credentials for both My Drive and the configured shared-drive lab.
- [ ] Regenerate the support matrix and fail review if it changes unexpectedly.
- [ ] Confirm public API/package compatibility against the intended breaking preview migration.
- [ ] Review current official Docs, Sheets, Slides, Drive, and OAuth release notes for schema drift.

## Official platform references

- [Google Docs tabs](https://developers.google.com/workspace/docs/api/how-tos/tabs)
- [Google Docs batchUpdate and WriteControl](https://developers.google.com/workspace/docs/api/reference/rest/v1/documents/batchUpdate)
- [Google Sheets batch updates](https://developers.google.com/workspace/sheets/api/guides/batchupdate)
- [Google Sheets developer metadata](https://developers.google.com/workspace/sheets/api/guides/metadata)
- [Google Slides batchUpdate](https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/batchUpdate)
- [Google Slides speaker notes](https://developers.google.com/workspace/slides/api/guides/notes)
- [Google Drive uploads](https://developers.google.com/workspace/drive/api/guides/manage-uploads)
- [Google Drive downloads and exports](https://developers.google.com/workspace/drive/api/guides/manage-downloads)
- [Google Drive comments](https://developers.google.com/workspace/drive/api/guides/manage-comments)
- [Google Drive shared drives](https://developers.google.com/workspace/drive/api/guides/manage-shareddrives)
- [Google Drive change tracking](https://developers.google.com/workspace/drive/api/guides/about-changes)
- [Google Workspace authentication](https://developers.google.com/workspace/guides/auth-overview)
- [OAuth installed applications](https://developers.google.com/identity/protocols/oauth2/native-app)

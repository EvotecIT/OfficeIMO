# OfficeIMO Reader Package Ownership

This document records the selective Reader package contract. The former `OfficeIMO.Reader` convenience package is
intentionally retired; there is no compatibility/meta package with that identity.

## Package model

```text
OfficeIMO.Reader.Core
├── contracts, schemas, normalized results, diagnostics, and limits
├── registry, neutral detection, orchestration, processors, and nested delegation
└── no OfficeIMO format-engine dependency

OfficeIMO.Reader.Word       -> Core + OfficeIMO.Word
OfficeIMO.Reader.Excel      -> Core + OfficeIMO.Excel + OfficeIMO.CSV
OfficeIMO.Reader.PowerPoint -> Core + OfficeIMO.PowerPoint
OfficeIMO.Reader.Markdown   -> Core + OfficeIMO.Markdown
OfficeIMO.Reader.Email      -> Core + OfficeIMO.Email
OfficeIMO.Reader.Pdf        -> Core + OfficeIMO.Pdf
OfficeIMO.Reader.*          -> Core + the owning format engine or explicit provider

OfficeIMO.Reader.All        -> composition only; every local managed adapter
```

Public namespaces remain `OfficeIMO.Reader`; `.Core` describes package and assembly ownership, not a namespace
migration. Applications install Core plus only the formats they need. Applications that intentionally want every
local managed handler install All. OCR engines, external processes, network clients, hosted providers, and native
tools remain explicit host choices.

## Email decision

`OfficeIMO.Reader.Email` is one adapter over Core and the unified `OfficeIMO.Email` package. It registers individual
artifacts, calendars/cards, PST/OST/OLM/EMLX and mailbox-directory stores, and OAB data. Store and AddressBook remain
separate semantic API areas inside `OfficeIMO.Email`, but they do not create separate Reader packages or dependency
layers.

## Completed contracts

- [x] Keep Core free of OfficeIMO format-engine project/package references.
- [x] Make `OfficeDocumentReader.Default` an empty immutable reader rather than a hidden convenience graph.
- [x] Keep registration instance-scoped through `OfficeDocumentReaderBuilder`.
- [x] Capture format options defensively when handlers are registered.
- [x] Support chunk, rich-result, and native async path/stream delegates.
- [x] Share bounded behavior across path, stream, byte, non-seekable, folder, and batch inputs.
- [x] Preserve neutral compound detection without pulling Email or document engines into Core.
- [x] Delegate nested archive/mailbox/attachment content only to handlers configured by the host.
- [x] Identify handler origin explicitly as `OfficeIMO` or `Custom`; do not retain the obsolete `IsBuiltIn` model.
- [x] Consolidate individual Email, Store, and OAB projection into `OfficeIMO.Reader.Email`.
- [x] Add selective Word, Excel, PowerPoint, and Markdown packages and remove their implementations from Core.
- [x] Keep All composition-only and exclude OCR/process/network/provider packages.
- [x] Validate every packed dependency group and clean-consumer install before release.
- [x] Include the breaking-package migration notes in this document and the final release-ready pull request.

## Ownership rules

1. Reusable parsing and inspection behavior belongs to the owning format package.
2. A Reader adapter owns registration, option translation, source mapping, and shared-model projection only.
3. `ReaderInputLimits` governs file, byte, and stream bounds.
4. Caller-owned stream lifetime and position promises are preserved.
5. Limits, unsupported content, and recoverable failures use stable structured diagnostics.
6. Optional native, platform, cloud, process, and provider dependencies stay outside Core and All unless a package
   name explicitly promises that dependency.
7. A new package needs a real dependency or runtime boundary; namespaces alone do not justify a NuGet layer.

## Release gates

- [x] Run the full Reader suite after the final package/API cleanup.
- [x] Pack Core, Email, Word, Excel, PowerPoint, Markdown, PDF, and All for every supported target group.
- [x] Inspect the actual `.nuspec` dependency groups, not only project references.
- [x] Prove a clean consumer can install Core plus one adapter without unrelated format engines.
- [x] Prove All registers the documented complete local managed handler set.
- [x] Verify the root package map, build configuration, website API paths, and adapter READMEs together.

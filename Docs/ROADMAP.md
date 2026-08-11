# OfficeIMO roadmap

This is the repository's single product backlog. It contains open work only. Implemented behavior is documented in package READMEs, support matrices, generated inventories, and current-state guides linked from [the documentation index](README.md).

An item belongs here when it has a clear product outcome and an owning package. Implementation checkpoints, completed task lists, release-wait notes, architectural rules, and competitor parity tables do not belong here.

Deliberately bounded compatibility contracts are not backlog by themselves. A preserved or rejected profile with a documented diagnostic becomes roadmap work only when the repository adopts a concrete supported shape and evidence plan for it.

## Release-wide quality

- [ ] Extend the generated Office compatibility catalog beyond the current Word, Excel, and PowerPoint legacy-format families into a package-neutral operation model for create, read, edit, preserve, inspect, convert, export, reject, and unsupported behavior.
- [ ] Generate compatible package README sections, website capability pages, MCP discovery, and support matrices from that model wherever one source can truthfully own the claim.
- [ ] Expand cross-producer fixture corpora with producer/version provenance and stable package or semantic diff policies.
- [ ] Add reproducible correctness, file-size, elapsed-time, peak-memory, allocation, cancellation, and deterministic-output evidence for representative workloads on every supported operating system.
- [ ] Add shared conversion reports and strict no-loss policies wherever an adapter can simplify, omit, rasterize, or preserve unsupported content.

## PDF engine and conversion fidelity

- [ ] Complete managed rendering for difficult Type 3 glyph programs, ICC mBA output transforms, higher-channel profiles, color-managed JPEG/DCT sample paths, higher-dimensional alternate-color transforms beyond the bounded Type 0/2/3/4 function pipeline, advanced pattern resources, transparency groups, masks, and optional-content interactions. Require producer fixtures, per-page diagnostics, and raster plus semantic proof before removing a limitation.
- [ ] Expand native Word, Excel, and PowerPoint reference corpora around DrawingML, SmartArt, chart variants, floating and grouped objects, headers and footers, print layout, and theme inheritance. Keep each improvement in the shared Drawing or PDF owner when more than one adapter needs it.
- [ ] Add exact-artifact external validation lanes for every supported PDF/A-2, PDF/A-3, PDF/A-4, PDF/UA-1, PDF/UA-2, Factur-X, and ZUGFeRD profile across representative generated artifacts and at least one independent producer corpus.
- [ ] Extend PDF-to-Office hybrid reconstruction beyond detected tables with bounded positioned text, image, and vector layers that remain visibly aligned while reports distinguish editable, visual-only, approximated, and omitted content.
- [ ] Publish optional provider adapters for OCR and other heavyweight engines only where the provider contract, resource limits, provenance, and warning merge are stable; keep `OfficeIMO.Pdf` dependency-light.
- [ ] Evaluate an explicit XFA inspection or conversion product only with licensed specification coverage, hostile-input limits, external fixtures, and a fail-closed migration path to AcroForm or static visual output. Do not execute XFA in the core reader.
- [ ] Expand the static HTML-to-PDF standards corpus for fragmentation, paged media, typography, SVG effects, and resource policy. Keep browser JavaScript and interactive state outside the managed static-renderer claim unless a separate browser-backed adapter owns them.

## Security and protected content

- [ ] Add interoperable ODF encryption/decryption only after an external producer corpus, explicit password and key policy, and fail-safe preservation evidence are available.

## Document-format depth

- [ ] Extend Excel/ODS conversion beyond the current typed formula, value, annotation, number-format, and validation subsets to date/time/custom-formula validations, conditional formatting, charts, pivot tables, and producer-backed advanced style fidelity.
- [ ] Add recursive typed ODT and ODP inline syntax with inherited-style resolution so nested spans and hyperlinks can convert without the current explicit flattening approximation.
- [ ] Expand Word/ODT and PowerPoint/ODP conversion for fields, notes, section/master-specific layout, advanced drawing geometry, media, and animation timing with producer fixtures and strict loss evidence.

## Completion rule

Remove an item when its public API, compatibility boundary, tests, generated evidence, and user documentation agree. GitHub Releases records delivered history, while `MIGRATION.md` retains only upgrade actions; this file does not retain completed milestones.

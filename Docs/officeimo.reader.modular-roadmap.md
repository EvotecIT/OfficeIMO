# OfficeIMO Reader Package Ownership

This document records the current Reader dependency problem and the target package boundary. The extraction should be
implemented as a dedicated change stacked after the unified `OfficeIMO.Email` package work; moving Reader contracts
while the Email production projects are also being consolidated would make both migrations harder to review.

## Current package shape

`OfficeIMO.Reader` is currently both the orchestration contract and a convenience bundle. Its packed `2.0.1` NuGet
declares these dependencies on every target:

```text
OfficeIMO.Reader
├── OfficeIMO.Drawing
├── OfficeIMO.Email
├── OfficeIMO.Excel
├── OfficeIMO.Markdown
├── OfficeIMO.Pdf
├── OfficeIMO.PowerPoint
├── OfficeIMO.Word.Markdown
└── OfficeIMO.Word
```

`System.Text.Json` is added for `netstandard2.0` and `net472`. This means a host cannot currently buy only the Reader
contracts plus one format. `OfficeIMO.Reader.Pdf`, for example, depends on `OfficeIMO.Reader` and `OfficeIMO.Pdf`, even
though the former already brings PDF into the graph.

Some apparent overlap is intentional. The CSV, JSON, XML, and YAML adapters replace the dependency-free plain-text
fallback with structured projection. The PDF overlap is different: it is a packaging artifact caused by the root
Reader also owning a built-in PDF handler.

`OfficeIMO.Reader.All` is a packable composition project that references all managed local adapters, but it is not
currently published on NuGet. Several adapters are published individually; the Email Store and AddressBook adapters
are not. Package names, source projects, and publication state therefore do not currently communicate one consistent
selection model.

## Target package shape

```text
OfficeIMO.Reader.Core
├── contracts, schemas, results, diagnostics, limits
├── registry, detection, orchestration, processors
└── no OfficeIMO format-engine dependency

OfficeIMO.Reader.Email      -> Reader.Core + OfficeIMO.Email
OfficeIMO.Reader.Pdf        -> Reader.Core + OfficeIMO.Pdf
OfficeIMO.Reader.Word       -> Reader.Core + owning Word/Markdown projection
OfficeIMO.Reader.Excel      -> Reader.Core + OfficeIMO.Excel
OfficeIMO.Reader.PowerPoint -> Reader.Core + OfficeIMO.PowerPoint
OfficeIMO.Reader.Markdown   -> Reader.Core + OfficeIMO.Markdown
OfficeIMO.Reader.*          -> Reader.Core + the one owning engine or deliberate provider

OfficeIMO.Reader            -> convenience/meta package for the documented common set
OfficeIMO.Reader.All        -> every managed local adapter, deliberately large graph
```

The public namespaces can remain `OfficeIMO.Reader`; the new Core name describes package and assembly ownership, not
a namespace migration.

### Email decision

Create one `OfficeIMO.Reader.Email` adapter over `OfficeIMO.Reader.Core` and the unified `OfficeIMO.Email` package. It
should register individual artifacts, calendars/cards, PST/OST/OLM/EMLX stores, and OAB data through the one Email
owner. Do not publish new `Reader.EmailStore` and `Reader.EmailAddressBook` packages; their source can move into
`Reader.Email` because they add no separate dependency or runtime boundary.

Until Core exists, the current two adapter projects remain honest thin packages over `OfficeIMO.Reader` and
`OfficeIMO.Email`. Making `Reader.Email` depend on today's heavy `OfficeIMO.Reader` would only add another package name
without reducing a single dependency.

### Convenience and All

- `OfficeIMO.Reader.Core` is the selection point for hosts that want the contracts and one or two adapters.
- `OfficeIMO.Reader` preserves the easy existing install and default common-format experience. Its intentionally broad
  dependency graph must be documented as a convenience choice.
- `OfficeIMO.Reader.All` is an explicit opt-in to every managed local adapter. It must be published and tested as a
  real meta package or removed; keeping an unpublished package-shaped project is misleading.
- OCR engines, external processes, network transports, and host-selected providers remain outside `Reader.All`.

## Completed modularization contracts

- `OfficeDocumentReader.Default` remains the convenience instance; modular registration stays instance-scoped.
- `OfficeDocumentReaderBuilder` freezes handlers, options, concurrency, and processors into an isolated reader.
- Adapters expose matching builder extensions such as `AddPdfHandler()` and capture defensive option snapshots.
- Registrations can provide chunk, native rich-result, and native asynchronous path/stream delegates.
- Path, stream, byte, and non-seekable inputs share the same bounded behavior.
- Capability manifests distinguish chunk, rich-result, and native async support.
- Ordered processors, bounded structured extraction, token-aware hierarchical chunking, and optional OCR build on the
  shared result instead of adding format-specific host pipelines.
- Web retrieval remains opt-in, uses a caller-owned `HttpClient`, and is not composed by `Reader.All`.

## Extraction order

- [ ] Freeze the public Reader contract and adapter registration surface with contract-focused tests.
- [ ] Create `OfficeIMO.Reader.Core` and move only contracts, schemas, orchestration, limits, processors, and neutral
  detection into it.
- [ ] Move each built-in format projection into a `Reader.*` adapter with exactly one owning engine dependency.
- [ ] Combine individual Email, Store, and OAB projection into `OfficeIMO.Reader.Email`; remove the unpublished split
  Email adapter projects.
- [ ] Convert `OfficeIMO.Reader` into the documented convenience/meta package without duplicating implementations.
- [ ] Make `OfficeIMO.Reader.All` the complete managed composition package and decide publication explicitly.
- [ ] Validate minimal and convenience dependency graphs from packed artifacts on all target frameworks.
- [ ] Add migration notes covering package selection, default registration, and any assembly-name compatibility impact.

## Ownership rules for adapters

1. Put reusable parsing and inspection behavior in the owning format package.
2. Keep the Reader adapter to registration, option translation, source mapping, and shared-model projection.
3. Use `ReaderInputLimits` for file, byte, and stream bounds.
4. Preserve caller-owned stream lifetime and position where promised.
5. Emit stable structured diagnostics for limits, unsupported content, and recoverable failures.
6. Add instance-builder registration alongside any retained static compatibility registration.
7. Keep optional native, platform, cloud, or process dependencies outside Core and convenience packages unless the
   package name explicitly promises that provider.

## Release gates

Before publishing the new graph:

- cover path, stream, byte, async, cancellation, limits, malformed input, and deterministic output where applicable;
- validate source IDs, chunk IDs and hashes, locations, and rich-result relationships;
- build every supported target and inspect the actual packed `.nuspec` dependency groups;
- prove a clean consumer can install `Reader.Core` plus one adapter without unrelated format engines;
- prove `Reader` and `Reader.All` composition behavior from packed artifacts;
- update adapter, root Reader, package-map, and migration documentation together.

# OfficeIMO provenance support matrix

OfficeIMO separates bounded structural inspection and selective removal from optional cryptographic verification and signing. `OfficeIMO.Core` owns format parsers, transformation policy, Unicode evidence, and provider-neutral contracts. `OfficeIMO.Provenance.C2pa` is the optional adapter for verification and signing through a host-supplied `c2patool` executable.

## Evidence model

| Evidence | Meaning | What it does not prove |
| --- | --- | --- |
| Structural C2PA carrier | A recognized embedded manifest or external reference is present | Signature validity, signer identity, trust, or AI authorship |
| Verified Content Credential | The configured verifier validated content binding and signature data under its trust policy | That every statement is true or that unsigned content is human-authored |
| IPTC Digital Source Type | A producer declared a capture, algorithmic, trained-algorithmic, or composite source type | Independent detection of how the content was made |
| Provider signal | A named detector reported its own durable watermark, statistical text signal, visible disclosure, or deterministic artifact | A portable signal shared by other providers, or a universal AI verdict |
| Unicode integrity finding | An exact invisible or context-sensitive code point occurs at a reported offset | That the character is malicious, a watermark, or AI-generated |

## Core file carriers

| Format | C2PA carrier | IPTC Digital Source Type | Removal behavior |
| --- | --- | --- | --- |
| JPEG | Ordered APP11 JUMBF segments | Standard and Extended XMP APP1 packets | Removes an unambiguous, structurally valid manifest; removes only AI-source declarations from valid XMP |
| PNG | `caBX` before the contiguous `IDAT` sequence | UTF-8 XMP `iTXt` | Preserves invalid CRC, duplicate, misplaced, or structurally ambiguous carriers by default |
| WebP | Final `C2PA` chunk in a valid extended RIFF container | Advertised `XMP ` chunk after image payloads | Rewrites the RIFF size and preserves unrelated chunks |
| GIF89a | `C2PA_GIF` application extension | XMP application extension | Requires one complete image and an exact trailer; GIF87a provenance applications are preserved as unsupported carriers |
| TIFF / BigTIFF | Primary-IFD C2PA tag | XMP tag 700 | Preserves overlapping IFD, pixel, strip, tile, JPEG, and shared-value storage |
| SVG | `c2pa:manifest` text in SVG metadata | Metadata-scoped `x:xmpmeta` or direct RDF/IPTC scope | Rewrites bounded XML only when the selected carrier is structurally unambiguous |
| ZIP / OPC image packages | Native `META-INF/content_credential.c2pa` plus supported embedded images | Supported embedded-image XMP | Generic removal blocks signed packages; document owners must explicitly handle signature invalidation |
| Structured text | Legacy multiline delimited blocks plus the C2PA 2.4 same-line comment form whose payload is a manifest reference | Not applicable | Accepts exact common language comment envelopes, including `#`, `//`, `--`, `;`, `%`, `'`, `REM`, `::`, `/* */`, `<!-- -->`, and `<# #>`; rejects duplicate carriers and preserves bare same-line delimiters and other lookalikes |
| Variation-selector text | Encoded C2PA wrapper | Not applicable | Removes only complete, bounded wrappers |

`RequireStructurallyValidCarrier` defaults to `true`. Turning it off permits best-effort removal from malformed carriers and should be reserved for explicitly destructive cleanup workflows.

## Verification and provider boundary

Structural inspection reports whether the carrier shape is safe to interpret or mutate; it does not establish authenticity, content binding, signer identity, or certificate trust. `C2paToolProvenanceVerifier` in `OfficeIMO.Provenance.C2pa` performs optional provider-backed verification and fails closed on malformed provider reports. The external executable and trust material are supplied by the host and are not bundled with OfficeIMO.

`IOfficeProvenanceSignalDetector` is the extension point for vendor-specific watermark and disclosure services. Each result retains the provider name, signal type, and `Detected`, `NotDetected`, `Inconclusive`, `ProviderUnavailable`, or `Error` status. `OfficeProvenanceAssessment` combines those results with structural, verification, and Unicode evidence without producing an `IsAi` property.

Providers can opt into cooperative cancellation through `ICancellableOfficeProvenanceSignalDetector` and `ICancellableOfficeProvenanceVerifier`. The original interfaces remain supported, so existing provider implementations do not need to change.

## Transformation and authoring

`OfficeProvenanceLifecycle` compares the source and candidate bytes before commit. Its default `PreserveIfUnchanged` policy blocks changed output when the source has an embedded or external Content Credential. `RemoveInvalidated` records before/after evidence; `SignAsDerived` requires an `IOfficeProvenanceSigner`, signs immutable snapshots of the inspected source and candidate, and always passes the source snapshot as the parent ingredient. The lifecycle independently checks provider identity, output location, file existence, and actual committed structural evidence.

`C2paToolProvenanceSigner` requires c2patool 0.27.0 or newer, writes application-controlled claim-generator information, title, and actions, delegates key operations to the current signer subprocess protocol, confirms that the staged output contains an embedded manifest, and commits atomically. New claims begin with a concrete `Created` intent; parent-derived claims begin with `Opened`, and c2patool creates the matching `parentOf` ingredient reference. Later actions use `c2pa.actions.v2`. Its built-in c2patool credential path is explicit and development-only. The signer intentionally emits no watermark action because those actions require matching soft-binding assertions that OfficeIMO does not currently create.

For signed Office, OpenDocument, EPUB, or PDF packages, use the owning package's removal adapter so invalidated package signatures are blocked or removed under the caller's explicit policy. The generic lifecycle is suitable when generic carrier removal is sufficient or when a custom signer supports the target format.

## Text integrity

`OfficeTextIntegrityInspector` reports exact BOMs, zero-width characters, word joiners, bidi controls, Unicode tags, variation selectors, typographic spaces, selected invisible format characters, controls, and unpaired surrogates. Findings retain UTF-16 offsets and distinguish informational, context-dependent, and potentially dangerous values. `OfficeTextIntegrityCleaner` removes only findings explicitly selected by the caller and verifies that the selected code point still occupies the recorded range.

## Workflow and command-line orchestration

`OfficeIMO.Workflows` exposes `IOfficeProvenanceWorkflowRunner` and `OfficeProvenanceWorkflowCatalog` for desktop, PowerShell, command-line, and service hosts. Inspect and assess are report-only operations. Removal preserves the source format, calls the package-owned adapter, stages the result beside its destination, reopens it through the same owner, compares the reopened structural evidence with the removal result, and publishes it only after that check succeeds. Batches are sequential and have an explicit item limit.

Files with unknown extensions remain signature-driven: recognized image, HTML, and PDF content can use the corresponding removal path, while a generic ZIP package remains inspection-only until a document owner establishes its mutation contract. Batch removal rejects duplicate effective output paths before processing any item, including collisions caused by equal basenames in one output directory.

`OfficeIMO.Tool` exposes the same contract under `officeimo provenance`. Its default JSON envelopes carry versioned schema identifiers, while `--format text` provides an interactive summary. Structural findings do not by themselves produce a failing exit code. Invalid requests, missing or unsupported inputs, output failures, cancellation, and execution failures use the tool's stable shared exit codes.

Provider-backed verification and signal detection are dependency-injected workflow services. The default CLI does not bundle credentials, trust material, vendor APIs, or a `c2patool` executable, so its `assess` command reports structural and text-integrity evidence unless a host composes additional providers.

## Package-owned adapters

| Package | Inspected content | Signature policy during removal |
| --- | --- | --- |
| `OfficeIMO.Word` | DOCX, DOCM, DOTX, DOTM and supported embedded images | Blocks save by default; `RemoveInvalidatedSignatures` removes the owned OPC signature graph and application signature metadata |
| `OfficeIMO.Excel` | XLSX, XLSM, XLTX, XLTM, XLAM, XLSB and supported embedded images | Uses the same explicit policy; XLSB ownership and signature metadata are validated without opening it as SpreadsheetML |
| `OfficeIMO.PowerPoint` | PPTX, PPTM, POTX, POTM, PPSX, PPSM, PPAM and supported embedded images | Removes only relationship- and content-type-owned signature parts when explicitly requested |
| `OfficeIMO.Visio` | VSDX-family OPC packages and supported embedded images | Resolves Visio application metadata and signature relationships through the package graph before cleanup |
| `OfficeIMO.Html` | Native manifest scripts/links, recursive `iframe srcdoc`, active CSS image carriers, responsive image carriers, and embedded supported images | File APIs preserve the document encoding; the string API returns UTF-8 and normalizes only exact charset declarations |
| `OfficeIMO.Markdown` | Structured-text carriers | String APIs require valid Unicode; file APIs accept and preserve strict UTF-8 or BOM-marked UTF-16 LE/BE and UTF-32 LE/BE; preserves unrelated Markdown text |
| `OfficeIMO.OpenDocument` | Native ODF manifest entry and supported package images | Inspection accepts encrypted packages; mutation requires valid ODF identity and an unencrypted package, and explicit cleanup removes owned signature entries and manifest declarations |
| `OfficeIMO.Epub` | Native EPUB manifest entry and supported package images | Requires valid EPUB identity and OPF ownership; removal detects native signatures and explicit cleanup removes only EPUB-owned signature entries |
| `OfficeIMO.Pdf` | C2PA manifest attachments with `application/c2pa` and `/AFRelationship /C2PA_Manifest` | Rewrites bounded object graphs only when the attachment has no active structural role; signed, encrypted, ambiguous, or unsupported-filter documents are preserved |

## Known limits

- OfficeIMO does not alter visible pixels, reconstruct images, suppress durable media watermarks, or rewrite generated text to defeat statistical watermarking.
- OfficeIMO does not ship vendor detector credentials or private APIs. Applications can add authorized detectors through `IOfficeProvenanceSignalDetector`.
- Strict removal preserves duplicate, competing, malformed, unsupported, signed, or structurally shared carriers when a targeted rewrite could discard unrelated data.
- Provider-backed HTML assessment snapshots local relative external manifests. An absolute `file:` base is rejected before providers run because the unchanged snapshot would otherwise resolve that dependency back to mutable source storage.
- Resource limits are configurable through `OfficeProvenanceOptions` and `OfficeProvenanceRemovalOptions`; inputs that exceed them are rejected instead of partially inspected.
- Cryptographic signing and verification currently depend on the external `c2patool` adapter. Core inspection, assessment, lifecycle policy, text integrity, and removal remain dependency-free.
- Unix process containment requires `setsid` from `util-linux`. The adapter recognizes standard Linux paths and Homebrew's macOS keg paths, and fails closed before launching `c2patool` when a session launcher is unavailable.
- BMFF/AVIF, audio, and video C2PA carriers are not parsed by OfficeIMO.Core. The signer commits only formats that Core can independently confirm after the installed tool signs them.

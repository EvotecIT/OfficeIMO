# OfficeIMO provenance support matrix

OfficeIMO separates bounded structural inspection and selective removal from optional cryptographic verification. `OfficeIMO.Core` owns the format parsers and generic `OfficeProvenanceInspector` / `OfficeProvenanceRemover` APIs. `OfficeIMO.Security` owns verification that calls a host-supplied `c2patool` executable.

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
| Structured text | Delimited manifest block whose payload may be a `data:application/c2pa` URI | Not applicable | Preserves surrounding text and source ordering |
| Variation-selector text | Encoded C2PA wrapper | Not applicable | Removes only complete, bounded wrappers |

`RequireStructurallyValidCarrier` defaults to `true`. Turning it off permits best-effort removal from malformed carriers and should be reserved for explicitly destructive cleanup workflows.

## Verification boundary

Structural inspection reports whether the carrier shape is safe to interpret or mutate; it does not establish authenticity, content binding, signer identity, or certificate trust. `C2paToolProvenanceVerifier` in `OfficeIMO.Security` performs optional provider-backed verification and fails closed on malformed provider reports. The external executable and trust material are supplied by the host and are not bundled with OfficeIMO.

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

## Known limits

- OfficeIMO removes provenance metadata; it does not alter visible pixels, reconstruct images, or attempt model-specific watermark suppression.
- Strict removal preserves duplicate, competing, malformed, unsupported, signed, or structurally shared carriers when a targeted rewrite could discard unrelated data.
- Resource limits are configurable through `OfficeProvenanceOptions` and `OfficeProvenanceRemovalOptions`; inputs that exceed them are rejected instead of partially inspected.
- Cryptographic verification currently depends on the external `c2patool` adapter. Core inspection and removal remain dependency-free.

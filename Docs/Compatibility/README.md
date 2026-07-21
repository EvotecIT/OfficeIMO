# Office binary and modern-format compatibility

OfficeIMO treats compatibility as a feature-level contract, not as a claim that every file with a supported extension is interchangeable. A conversion either writes a native or equivalent representation, records an editable or visual fallback, retains the original source, or blocks. It does not silently discard a known feature.

The generated contracts in [`generated`](generated/README.md) are the source of truth for current coverage. They include the concrete format inventory and separate import, authoring, round-trip, modern-to-legacy, and legacy-to-modern states for every tracked capability.

## Implemented phases

- [x] Phase 0: shared format descriptors, feature states, impacts, reports, JSON, and Markdown catalogs
- [x] Phase 1: content-aware routing across Word, Excel, PowerPoint, and OfficeIMO.Reader
- [x] Phase 2: checked corpus manifests, hashes, bidirectional conversion cases, LibreOffice checks, and optional Microsoft Office desktop checks
- [x] Phase 3: Word DOC/DOT native conversion plus omission-free page-image and retained-source fallbacks
- [x] Phase 4: Excel XLS/XLSB native conversion plus palette-quantized worksheet visual and retained-source fallbacks
- [x] Phase 5: PowerPoint PPT/POT/PPS native conversion plus shape-level static visual and retained-source fallbacks

## Format coverage

| Family | Legacy binary | Modern |
| --- | --- | --- |
| Word | `.doc`, `.dot` | `.docx`, `.docm`, `.dotx`, `.dotm` |
| Excel | `.xls`, `.xlt`, `.xla`, `.xlm`, `.xlw` | `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xlam`, `.xlsb` |
| PowerPoint | `.ppt`, `.pot`, `.pps`, `.ppa` | `.pptx`, `.pptm`, `.potx`, `.potm`, `.ppsx`, `.ppsm`, `.ppam` |

Classification and import do not imply that every extension is a writable destination. The structured preflight report currently blocks legacy Excel `.xlt`, `.xla`, `.xlm`, and `.xlw` authoring and PowerPoint `.ppa` authoring. This distinction is represented in the catalogs and returned diagnostics.

## Conversion policy

`OfficeCompatibilityMode` makes the caller's preference explicit:

| Mode | Accepted representation |
| --- | --- |
| `StrictNative` | Native or semantically equivalent content only |
| `PreferEditable` | Native, equivalent, or a documented editable approximation; static-only substitutions block |
| `PreferVisual` | Preserves appearance with a static visual when editable representation is unavailable |
| `BestEffort` | Uses the safest available native, editable, visual, or deliberate-drop path and reports every loss |
| `PreservationOnly` | Permits fallbacks and embeds the complete original source for byte-level recovery |

The older family-specific `LossPolicy = Allow` setting maps to `BestEffort` when no compatibility mode was selected. New callers should prefer `CompatibilityMode` because it distinguishes editability from visual fidelity and preservation.

Use `AnalyzeConversion` before writing when an application needs approval or workflow routing. `Convert` returns the same immutable assessment together with the committed path. `RequireNoLoss()` rejects reports containing loss or blocked features.

```csharp
PowerPointPresentationConversionReport assessment =
    PowerPointPresentation.AnalyzeConversion(
        "source.pptx",
        "destination.ppt",
        new PowerPointPresentationConversionOptions {
            CompatibilityMode = OfficeCompatibilityMode.PreferVisual
        });

PowerPointPresentationConversionResult result =
    PowerPointPresentation.Convert(
        "source.pptx",
        "destination.ppt",
        new PowerPointPresentationConversionOptions {
            CompatibilityMode = OfficeCompatibilityMode.PreferVisual
        });
```

Analysis does not create or replace an artifact. Conversion writes to a staging file and commits atomically after the selected policy succeeds.

## Real fallback artifacts

Fallback states describe content that is actually written:

- Word renders every page with omission and failure checks, then writes the rendered pages as full-page DOC/DOT pictures.
- Excel renders each worksheet and writes an XLS-compatible, palette-quantized cell raster. This keeps the fallback readable in binary XLS and XLSB without claiming formula or object editability.
- PowerPoint converts charts, SmartArt, and unsupported tables to deterministic PNG picture shapes while retaining slide bounds. `PreferEditable` blocks these substitutions because the result is static.

Every fallback reports its state, affected fidelity dimensions, source location when available, and generated artifact identity.

## Recovering an embedded source

Set the family conversion option `EmbedSourceWhenLossy` when byte-level recovery is required. `PreservationOnly` enables source retention automatically.

After loading the converted file, call the family document's `TryGetCompatibilitySourcePayload` method:

```csharp
using PowerPointPresentation presentation =
    PowerPointPresentation.Load("converted.ppt");

if (presentation.TryGetCompatibilitySourcePayload(
        out OfficeCompatibilitySourcePayload? source,
        out string? error)) {
    byte[] originalBytes = source!.ToArray();
    string verifiedSha256 = source.Sha256;
}
```

The carrier records the concrete source format, original file name, compatibility mode, and SHA-256 digest; the extraction API also exposes the verified payload length. Extraction verifies the digest before returning bytes and enforces a 512 MiB payload limit. The payload is inert to OfficeIMO, but it may contain macros, embedded objects, external links, hidden content, or other active material from the original file. Treat extracted bytes as untrusted input and apply the same security policy used for the source.

## Validation gates

The repository gate verifies generated contracts, corpus identity, import reports, conversion and reopen behavior, package security, preservation, and checked visual baselines:

```powershell
pwsh -NoProfile -File Build/Test-OfficeInteroperabilityGate.ps1 -Suite Full
```

LibreOffice provides an independent open-and-convert oracle for every checked DOC, XLS, XLSB, and PPT corpus artifact:

```powershell
pwsh -NoProfile -File Build/Test-OfficeCorpusLibreOffice.ps1 `
  -OutputDirectory /path/to/empty/output
```

By default, that lane additionally creates fresh OfficeIMO DOC, XLS, XLSB, and
PPT files from modern source documents, verifies their hashes, and asks
LibreOffice to open and convert each file; the generated PPT is also rendered
to PDF. This keeps the external oracle attached to current writer output rather
than only to historical corpus inputs.

On Windows with desktop Office installed, add `-MicrosoftOffice` to the interoperability gate to exercise Word, Excel, and PowerPoint COM oracles. These external applications are validation oracles; they are not runtime conversion dependencies.

When coverage changes, update the owning family catalog and tests, regenerate `generated`, and keep the report honest about any approximation, rasterization, retained carrier, omission, or block.

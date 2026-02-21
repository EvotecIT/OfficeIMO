# OfficeIMO Reader Modularization Roadmap (Internal Preview)

This plan introduces reusable format packages first, then composes them into Reader adapters.
The goal is clean IntelligenceX integration without forcing one large dependency set on all users.

## Principles

- Keep `OfficeIMO.Reader` as the stable ingestion facade.
- Move format-specific behavior into reusable packages.
- Keep heavy features optional and isolated.
- Preserve deterministic contracts (`ReaderChunk`, progress, warnings, source metadata).

## Package Structure (Scaffolded)

- `OfficeIMO.Zip`
- `OfficeIMO.Epub`
- `OfficeIMO.Reader.Zip`
- `OfficeIMO.Reader.Epub`
- `OfficeIMO.Reader.Text`
- `OfficeIMO.Reader.Html`

All scaffolded packages are currently excluded from publishing.

## Delivery Order (Low -> High Effort)

1. ZIP path (low)
- Harden archive traversal. (implemented in scaffold branch)
- Add entry-level warning handling. (implemented in scaffold branch)
- Support nested Office/text extraction from ZIP entries. (implemented in scaffold branch)
- Add stream ingestion + registry stream dispatch. (implemented in scaffold branch)

2. EPUB path (low-medium)
- Basic chapter extraction from XHTML/HTML entries.
- Normalize into Reader chunks.
- Add OPF/spine/nav-aware ordering in next pass. (implemented in scaffold branch)
- Add stream ingestion + registry stream dispatch. (implemented in scaffold branch)

3. Structured text path (medium)
- CSV semantic chunks with tables.
- JSON/XML structural chunkers. (implemented in scaffold branch)
- Stable markdown previews and table metadata.
- Add stream ingestion + registry stream dispatch. (implemented in scaffold branch)
- Reuse `OfficeIMO.CSV` stream API for CSV path/stream parity. (implemented in scaffold branch)

4. HTML path (medium)
- HTML -> Word -> Markdown bridge.
- Chunking and source/citation metadata.
- Performance and fidelity tuning for large inputs.
- Add stream ingestion + registry stream dispatch. (implemented in scaffold branch)
- Add configurable HTML/Markdown conversion options for adapter registration. (implemented in scaffold branch)

5. Reader core modularization (medium-high)
- Replace hardcoded switch routing with handler registry. (implemented in scaffold branch)
- Add capability discovery for host apps. (implemented in scaffold branch)
- Keep existing `DocumentReader.Read*` API behavior stable.

6. Optional heavy paths (high)
- OCR/image text extraction.
- Audio transcription.
- URL/Youtube ingestion.
- Keep each as separate opt-in package.

## IntelligenceX Integration Expectations

- Current `OfficeImoReadTool` contract remains stable.
- New formats become available incrementally via adapters.
- Missing optional dependencies should emit warnings, not hard failures.
- Capability listing can be surfaced to chat/tooling once handler registry lands.

## Release Gating

Before publishing any new package:

- Add dedicated tests (success, error, limits, deterministic output).
- Validate chunk IDs/source metadata stability.
- Verify no regressions for existing Reader consumers.
- Confirm package-level dependency footprint is intentional.

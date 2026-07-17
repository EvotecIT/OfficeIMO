# OfficeIMO Reader Package Ownership

This document records the current Reader package boundaries and the remaining modular follow-up work.

## Current package shape

| Owner | Responsibility |
| --- | --- |
| `OfficeIMO.Reader` | Shared chunks and rich result, built-in default instance, isolated reader instances, bounded sync/async execution, detection, diagnostics, processors, structured extraction, hierarchical chunking, and OCR execution contracts. |
| Format packages | Parse and inspect their own formats. Reader must not duplicate Word, Excel, PowerPoint, Markdown, PDF, RTF, HTML, EPUB, Visio, CSV, ZIP, or other format logic. Small transport formats without an owning package may use a narrow, isolated adapter. |
| `OfficeIMO.Reader.*` adapters | Register path/stream handlers and project owning format models into the shared Reader contracts. |
| `OfficeIMO.Reader.Web` | Explicit caller-injected HTTP transport that bounds retrieval and routes bytes through an existing Reader instance. |
| `OfficeIMO.Reader.Ocr.Process` | Optional versioned external-process OCR bridge. |
| `OfficeIMO.Reader.Ocr.Tesseract` | Optional Tesseract CLI provider with line and word geometry. |

The modular adapters are working packages in the publishing pipeline, not placeholders:

- `OfficeIMO.Reader.Csv`
- `OfficeIMO.Reader.Epub`
- `OfficeIMO.Reader.Html`
- `OfficeIMO.Reader.Image`
- `OfficeIMO.Reader.Json`
- `OfficeIMO.Reader.Notebook`
- `OfficeIMO.Reader.Pdf`
- `OfficeIMO.Reader.Rtf`
- `OfficeIMO.Reader.Subtitles`
- `OfficeIMO.Reader.Visio`
- `OfficeIMO.Reader.Web`
- `OfficeIMO.Reader.Xml`
- `OfficeIMO.Reader.Yaml`
- `OfficeIMO.Reader.Zip`
- `OfficeIMO.Reader.Ocr.Process`
- `OfficeIMO.Reader.Ocr.Tesseract`

## Completed modularization contracts

- `OfficeDocumentReader.Default` remains the built-in-only convenience instance; modular registration stays instance-scoped.
- `OfficeDocumentReaderBuilder` freezes handlers, options, concurrency, and processors into an isolated `OfficeDocumentReader`.
- Adapters expose matching builder extensions such as `AddPdfHandler()` and capture defensive option snapshots during registration.
- Registrations can provide chunk delegates, native rich-result delegates, and native asynchronous path/stream delegates.
- Path, stream, byte, and non-seekable input paths share the same bounded input behavior.
- Capability manifests distinguish chunk, rich-result, and native async support.
- The stable rich transport starts at `OfficeDocumentReadResult` schema version 5; current version 6 adds calendar and vCard kinds while the Reader continues to accept version 5 payloads.
- Ordered processors, bounded structured extraction, token-aware hierarchical chunking, and optional OCR build on the shared result instead of adding format-specific host pipelines.
- OCR providers remain opt-in and do not change the core package dependency graph.
- Web retrieval remains opt-in, uses a caller-owned `HttpClient`, and is not composed by `OfficeIMO.Reader.All` or the global tool.

## Ownership rules for new adapters

1. Put reusable parsing and inspection behavior in the owning format package.
2. Keep the Reader adapter to registration, option translation, source mapping, and shared-model projection.
3. Use `ReaderInputLimits` for file, byte, and stream bounds.
4. Preserve caller-owned stream lifetime and position where the contract promises it.
5. Emit stable structured diagnostics for limits, unsupported content, and recoverable failures.
6. Add instance-builder registration alongside any static compatibility registration.
7. Keep optional native, platform, cloud, or process dependencies outside `OfficeIMO.Reader`.

## Release gates

Before publishing a new or changed adapter:

- cover path, stream, byte, async, cancellation, limits, malformed input, and deterministic output where those surfaces apply;
- validate source IDs, chunk IDs/hashes, locations, and rich-result relationships;
- build every supported target framework and pack the affected public packages;
- verify that optional dependencies are intentional and do not leak into the core package;
- update the adapter README, the core Reader README, and the website Reader page when the public workflow changes.

## Follow-up candidates

- Continue increasing format fidelity in the owning packages and project that evidence through the adapters.
- Add optional providers only when their dependency and runtime boundaries are clear.
- Keep audio, video, and transcription ingestion in separate packages if downstream products require them; URL retrieval remains isolated in `OfficeIMO.Reader.Web`.
- Keep email and attachment ingestion in the existing Mailozaurr owner; it is outside the current OfficeIMO Reader stack.

# OfficeIMO.Word Review - 2026-03-08

Branch: `codex/word-review-20260308`

## Scope

- Reviewed `OfficeIMO.Word` implementation with focus on document lifecycle, stream/file handling, compatibility helpers, and public API consistency.
- Cross-checked the implementation against existing Word-focused tests in `OfficeIMO.Tests` and `OfficeIMO.VerifyTests`.
- Ran:
  - `dotnet build OfficeIMO.Word/OfficeIMO.Word.csproj -c Debug`
  - `dotnet test OfficeIMO.Tests/OfficeIMO.Tests.csproj -c Debug --filter "FullyQualifiedName~Word" --no-restore`

## Key Findings

### 1. `LoadAsync` has stricter file access than `Load`

`WordDocument.LoadAsync(...)` opens the source file with write access when `readOnly == false` and only `FileShare.Read`, even though it immediately copies the file into memory. The synchronous `Load(...)` path correctly uses read access and permissive sharing.

Impact:

- `LoadAsync` can fail on documents that `Load` successfully opens.
- Shared test fixtures and user workflows involving files opened by Word, Explorer previews, antivirus, or backup agents are more likely to fail on the async path.

Recommended fix:

- Mirror the synchronous open strategy in `LoadAsync`: use read access and `FileShare.ReadWrite | FileShare.Delete`.

### 2. Stream save paths skip the OpenOffice compatibility fixer

`Save(string)`, `SaveAs(string)`, and `SaveAsByteArray()` call `Helpers.MakeOpenOfficeCompatible(...)`, but `Save(Stream)` does not. That means `Save(Stream)`, `SaveAs(Stream)`, and `SaveAsMemoryStream()` can emit a different package shape than file-based saves.

Impact:

- Different save destinations can produce different compatibility outcomes for the same document.
- Any OpenOffice-specific relationship normalization the project relies on is bypassed for stream-based workflows.

Recommended fix:

- Run the same compatibility pass for stream saves before rewinding the output stream.
- Add a regression test comparing file save and stream save relationship parts.

### 3. `WordDocumentComparer.Compare` returns a document whose `FilePath` is already deleted

`Compare(...)` loads the result from a temp file and then deletes that temp file in `finally` before returning the `WordDocument`.

Impact:

- The returned document has a `FilePath` that no longer exists.
- `Open()` will fail, and `Save()` recreates the deleted temp file path unexpectedly.

Recommended fix:

- Either keep the temp file and document the ownership model, or clear `FilePath` before returning.
- Better: return an in-memory document explicitly detached from any ephemeral temp path.

### 4. `Helpers.Cloners` leaks file handles

`CopyFileStreamToMemoryStream` and `CopyFileStreamToFileStream` open source streams without disposing them.

Impact:

- Unnecessary file locks until GC/finalization.
- Risk of intermittent test failures and surprising behavior for callers that copy and then immediately mutate/delete source files.

Recommended fix:

- Wrap source streams in `using`.
- Add a small test asserting the source file is deletable immediately after the helper returns.

## Improvement Opportunities

### Architecture

- Split more `WordDocument` lifecycle code into dedicated services or internal helper types.
- Reduce public facade size in `WordDocument.cs`, `WordDocument.PublicMethods.cs`, `WordImage.cs`, and `WordParagraph.PublicMethods.cs`.
- Standardize sync/async parity reviews for every public lifecycle method.

### API Consistency

- Align image-download behavior between `WordDocument.AddImageFromUrl(...)` and `Fluent.ImageBuilder.AddFromUrlAsync(...)`.
- Consider shared download validation: content type checks, max size guards, cancellation support, and reusable `HttpClient`.
- Audit XML comments where behavior and implementation diverge, especially stream-based creation/save semantics.

### Tests

- Add direct tests for `LoadAsync` share/read-only behavior.
- Add stream-save compatibility tests covering `Save(Stream)` and `SaveAsMemoryStream()`.
- Add verify/OpenXml validation coverage for `WordImage`, `WordTextBox`, embedded documents, mail merge, and comparer output.
- Add focused tests for `WordDocumentComparer.Compare` returned-document semantics.

## Validation Notes

- `OfficeIMO.Word` builds cleanly.
- Word-focused test execution mostly passed for the core library, but the run had repeated failures in `OfficeIMO.Word.Pdf` scenarios because QuestPDF native dependencies were incompatible in the current build output.
- Final test summary from the filtered run:
  - Passed: `1013`
  - Failed: `111`
  - Skipped: `5`
  - Total: `1129`

## Suggested Next Steps

1. Fix the four lifecycle/IO findings above first.
2. Add targeted regression tests before broader refactors.
3. Clean/rebuild the PDF stack and normalize QuestPDF native dependency resolution so Word-focused CI runs are trustworthy again.
4. Follow with a second review pass focused on DOM-heavy areas: images, text boxes, embedded content, and comparer output.

# OfficeIMO.Excel URI Loading Proposal

## Summary

ExcelFast's URL support is useful, but it is transport glue in front of a local-file loader. OfficeIMO.Excel should treat remote workbook loading as an engine feature instead of pushing that concern into PSWriteOffice. The engine already owns workbook stream and reader semantics, so it is the right place to own HTTP policy, download safety, cleanup, and the transition into the existing Open XML load path.

This branch implements the first slice as a bounded, memory-backed remote loader. It deliberately does not add temp-file mode or remote save semantics.

The recommended public surface is:

```csharp
ExcelDocument.Load(Uri uri, ExcelHttpLoadOptions? httpOptions = null, bool readOnly = true, OpenSettings? openSettings = null)
ExcelDocument.LoadAsync(Uri uri, ExcelHttpLoadOptions? httpOptions = null, bool readOnly = true, OpenSettings? openSettings = null, CancellationToken cancellationToken = default)

ExcelDocumentReader.Open(Uri uri, ExcelReadOptions? readOptions = null, ExcelHttpLoadOptions? httpOptions = null)
ExcelDocumentReader.OpenAsync(Uri uri, ExcelReadOptions? readOptions = null, ExcelHttpLoadOptions? httpOptions = null, CancellationToken cancellationToken = default)
```

PSWriteOffice can then expose a thin wrapper:

```powershell
Import-OfficeExcel -Uri https://example.test/report.xlsx
Get-OfficeExcel -Uri https://example.test/report.xlsx
```

No PSWriteOffice downloader, temp-file policy, retry policy, or content validation should be needed.

## Current Fit

OfficeIMO.Excel already has:

- `ExcelDocument.Load(string)` and `LoadAsync(string)` for path-backed workbooks.
- `ExcelDocument.Load(Stream)` and `LoadAsync(Stream)` for caller-provided streams.
- `ExcelDocumentReader.Open(string)`, `Open(Stream)`, and `Open(byte[])` for read-only reader workflows.
- Existing byte materialization and package content-type normalization before `SpreadsheetDocument.Open(...)`.
- Cancellation-aware async helpers in the load and read paths.

That means the first implementation does not need a new workbook parser. It needs a careful HTTP fetch helper that produces a bounded, seekable workbook package and then calls the existing byte or stream loader.

## Proposed Types

```csharp
public sealed class ExcelHttpLoadOptions {
    public ExcelUriSchemePolicy SchemePolicy { get; set; } = ExcelUriSchemePolicy.HttpsOnly;
    public long MaxBytes { get; set; } = 100L * 1024L * 1024L;
    public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(100);
    public string? UserAgent { get; set; } = "OfficeIMO.Excel";
    public IDictionary<string, string> Headers { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    public bool ValidateZipHeader { get; set; } = true;
    public bool ValidateContentTypeWhenPresent { get; set; } = false;
    public ISet<string> AllowedContentTypes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel",
        "application/octet-stream"
    };
    public IProgress<ExcelHttpLoadProgress>? Progress { get; set; }
    public HttpClient? HttpClient { get; set; }
}

public enum ExcelUriSchemePolicy {
    HttpsOnly,
    HttpAndHttps
}

public readonly struct ExcelHttpLoadProgress {
    public long BytesRead { get; }
    public long? ContentLength { get; }
}
```

The implementation should clone option values at operation start so callers cannot mutate active requests midway.

## Default Policy

- `https` only by default. Plain `http` requires `SchemePolicy = HttpAndHttps`.
- Full response download before Open XML parsing. This matches current package-opening reality and avoids pretending the parser can stream XLSX row data directly from HTTP.
- `MaxBytes` enforced against `Content-Length` when present and against the running byte count while copying.
- `Timeout` and `CancellationToken` both respected.
- Caller headers and user-agent supported without adding a dependency on any auth framework.
- Optional content-type validation because real file hosts often return generic values. ZIP header validation is the stronger default check for `.xlsx`.
- Redirect handling can use the default `HttpClientHandler` behavior for the internal client. If a caller supplies `HttpClient`, their handler policy wins.
- Remote load is read-only by default. It should not imply remote save or upload semantics.

## Ownership And Cleanup

Remote loading should not leak temp files or hidden state.

For this first slice, use memory for all supported downloads and pass bytes into existing load/open methods. This keeps ownership simple and mirrors the current local file and stream loaders, which already materialize package bytes.

For a future slice, add temp-file support only if it reduces real memory pressure in the underlying Open XML path. If used:

- Create temp files under an explicit option directory or `Path.GetTempPath()`.
- Use a package-owned cleanup scope so files are deleted when `ExcelDocument` or `ExcelDocumentReader` is disposed.
- Never expose the temp path as `FilePath`, because remote load does not have local save-back semantics.
- Do not copy back to the temp file on dispose unless the caller explicitly saves to a chosen destination.

## Implementation Plan

1. Add `ExcelHttpLoadOptions`, `ExcelUriSchemePolicy`, and `ExcelHttpLoadProgress` under `OfficeIMO.Excel`.
2. Add an internal `ExcelHttpWorkbookLoader` that:
   - validates the URI scheme,
   - sends a GET request with `ResponseHeadersRead`,
   - checks success status,
   - enforces timeout, cancellation, and byte limit,
   - copies to memory,
   - validates ZIP magic bytes when enabled,
   - reports progress during copy.
3. Add `ExcelDocument.Load(Uri, ...)` and `LoadAsync(Uri, ...)`.
   - Do not expose `autoSave` for URI loads. Remote save belongs to an explicit upload API later, not this convenience loader.
   - Call `LoadFromByteArray(...)` with `filePath: null` and `preferFilePathOnFallback: false`.
4. Add `ExcelDocumentReader.Open(Uri, ...)` and `OpenAsync(Uri, ...)`.
   - Call the existing `Open(byte[], ExcelReadOptions?)` path.
5. Add unit tests with an in-process HTTP server or injectable `HttpClient` handler:
   - HTTPS-only default rejects `http`.
   - Opt-in `http` works.
   - headers/user-agent are sent.
   - max byte limit rejects oversized responses before and during copy.
   - cancellation is observed while downloading.
   - content-type validation is opt-in.
   - invalid ZIP header is rejected.
   - reader and document paths both open a valid downloaded workbook.

## PSWriteOffice Surface

After the engine API exists, PSWriteOffice should only map PowerShell parameters to OfficeIMO options:

- `-Uri` as a separate parameter set from `-Path`.
- `-AllowHttp` maps to `SchemePolicy = HttpAndHttps`.
- `-Header` maps to `ExcelHttpLoadOptions.Headers`.
- `-TimeoutSeconds`, `-MaximumBytes`, and `-UserAgent` map directly.
- Progress can be wired to PowerShell progress from `ExcelHttpLoadProgress`.

Anything beyond that, including auth handlers, retries, cache policy, temp-file cleanup, and package validation, should remain in OfficeIMO.Excel or be provided by caller-supplied `HttpClient`.

## Non-Goals

- No streaming XLSX parsing directly from HTTP in the first slice.
- No remote save/upload behavior.
- No cookie jar or browser-session semantics.
- No automatic credential discovery.
- No wrapper-side fallback downloader in PSWriteOffice.

## Future Questions

- Should a future temp-file mode be added after measuring whether it reduces real memory pressure in the current Open XML path?
- Should `MaxBytes` stay at `100 MB`, move to `256 MB`, or become unlimited with a documented recommendation? A bounded default is safer, but large operational workbooks may need an easy override.
- Should content-type validation stay opt-in forever, or become a warn-only diagnostic hook once OfficeIMO has a broader diagnostics surface?

# OfficeIMO.Excel remote loading

OfficeIMO.Excel owns remote workbook loading because HTTP policy, download safety, and package validation belong next to the workbook engine. Consumers should not duplicate download, temporary-file, redirect, or content-validation logic.

## Public API

Remote I/O is asynchronous only:

```csharp
using OfficeIMO.Excel;

Uri uri = new("https://example.test/report.xlsx");

using ExcelDocument workbook = await ExcelDocument.LoadAsync(
    uri,
    new ExcelHttpLoadOptions {
        MaxBytes = 100L * 1024L * 1024L
    },
    cancellationToken: cancellationToken);

using ExcelDocumentReader reader = await ExcelDocumentReader.OpenAsync(
    uri,
    cancellationToken: cancellationToken);
```

There are no synchronous `Load(Uri)` or `Open(Uri)` counterparts. A remote request cannot complete synchronously, and hiding it behind a synchronous API would create blocking and cancellation problems. Local path, stream, and byte-array inputs retain their synchronous entry points.

## Default policy

- HTTPS is required by default. Plain HTTP requires an explicit `ExcelUriSchemePolicy.HttpAndHttps` opt-in.
- The response is downloaded completely before Open XML parsing. XLSX packages require seekable ZIP access; the API does not pretend to stream rows directly from HTTP.
- `MaxBytes` is checked against `Content-Length`, when present, and against the bytes actually read.
- Timeout and caller cancellation are both observed.
- Redirects are followed manually so the target scheme is revalidated at every hop.
- Custom headers are removed after a redirect to another host.
- ZIP header validation is enabled by default. Content-type validation is optional because real file hosts often use generic MIME types.
- Remote loads are detached. They never imply upload or save-back behavior.

Options are snapshotted when the operation starts, so mutating the caller-owned options object does not alter an in-flight request.

## Lifecycle and ownership

The loader materializes a bounded, seekable package in memory and hands it to the same native Open XML load path used by byte-array input. No temporary file or hidden path becomes associated with the returned document.

The caller owns the returned `ExcelDocument` or `ExcelDocumentReader` and must dispose it. To persist a remote workbook locally, choose the destination explicitly:

```csharp
using ExcelDocument workbook = await ExcelDocument.LoadAsync(uri, cancellationToken: cancellationToken);
workbook.SaveCopy("report.xlsx");
```

`DocumentPersistenceMode.SaveOnDispose` is rejected for remote loads because a detached remote document has no writable destination. Remote upload is a separate concern and is not inferred from the source URI.

## Thin consumer surfaces

A PowerShell or CLI wrapper should expose an asynchronous command surface and map its parameters directly to `ExcelHttpLoadOptions`. The wrapper may present values such as URI, maximum bytes, timeout, headers, user agent, and HTTP opt-in, but it should not implement its own downloader, retry policy, redirect handling, ZIP validation, or temporary-file cleanup.

Authentication in this contract is explicit through request headers. Richer transport customization should be added only if OfficeIMO can continue enforcing redirect, scheme, byte-limit, and credential-forwarding policy.

## Non-goals

- Streaming XLSX row parsing directly from HTTP.
- Implicit remote save or upload.
- Browser-session or cookie-jar behavior.
- Automatic credential discovery.
- Consumer-side fallback downloaders.

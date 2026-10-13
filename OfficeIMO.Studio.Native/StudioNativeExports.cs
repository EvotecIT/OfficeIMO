using System.Runtime.InteropServices;
using System.Text.Json;
using OfficeIMO.Studio.Features.Conversion;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Scan;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Native;

/// <summary>
/// The C entry points. Each call blocks until the engine finishes, so hosts call from a background thread or task.
/// Fallible calls return a status code and, on failure, a message through <c>error</c> (may be null if the host does
/// not want it). Everything the library returns is released with <c>officeimo_free</c>; handles have their own release.
/// </summary>
internal static unsafe class StudioNativeExports {
    [UnmanagedCallersOnly(EntryPoint = "officeimo_abi_version")]
    internal static int AbiVersion() => StudioNativeAbi.Version;

    [UnmanagedCallersOnly(EntryPoint = "officeimo_library_version")]
    internal static byte* LibraryVersion() =>
        StudioNativeAbi.ToUtf8(typeof(StudioNativeExports).Assembly.GetName().Version?.ToString() ?? "0.0.0");

    [UnmanagedCallersOnly(EntryPoint = "officeimo_free")]
    internal static void Free(void* memory) => NativeMemory.Free(memory);

    /// <summary>Names the app-owned folder for recovery copies. Required once before opening documents.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_initialize")]
    internal static int Initialize(byte* dataDirectory, byte** error) {
        string? directory = StudioNativeAbi.FromUtf8(dataDirectory);
        return Run(error, 0, _ => StudioNativeEngine.Initialize(directory!));
    }

    // Cancellation -----------------------------------------------------------------------------------------------

    [UnmanagedCallersOnly(EntryPoint = "officeimo_cancellation_create")]
    internal static nint CancellationCreate() => GCHandle.ToIntPtr(GCHandle.Alloc(new CancellationTokenSource()));

    /// <summary>Requests cancellation of every call that received this token; safe from any thread, including during a call.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_cancellation_cancel")]
    internal static void CancellationCancel(nint cancellation) {
        if (Resolve<CancellationTokenSource>(cancellation) is { } source) source.Cancel();
    }

    [UnmanagedCallersOnly(EntryPoint = "officeimo_cancellation_free")]
    internal static void CancellationFree(nint cancellation) {
        if (cancellation == 0) return;
        GCHandle handle = GCHandle.FromIntPtr(cancellation);
        (handle.Target as CancellationTokenSource)?.Dispose();
        handle.Free();
    }

    /// <summary>
    /// Makes a PDF at <paramref name="outputPath"/> from scanned page images and the words the host recognized on them,
    /// described as UTF-8 JSON, and returns a summary as JSON.
    /// </summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_create_pdf_from_scan_json")]
    internal static int CreatePdfFromScanJson(byte* request, byte* outputPath, nint cancellation, byte** json, byte** error) {
        if (json is null) return Fail(error, new ArgumentNullException(nameof(json)));
        *json = null;
        StudioNativeScanRequest? parsed;
        try { parsed = JsonSerializer.Deserialize(StudioNativeAbi.FromUtf8(request) ?? "null", StudioNativeJson.Default.StudioNativeScanRequest); }
        catch (JsonException exception) { return Fail(error, new ArgumentException("The scan request is not valid JSON.", exception)); }
        if (parsed?.Pages is null) return Fail(error, new ArgumentNullException(nameof(request)));
        string? destination = StudioNativeAbi.FromUtf8(outputPath);
        return Run(error, cancellation, token => {
            var pages = parsed.Pages.Select(page => new PdfScanPage(page.Image ?? string.Empty,
                (page.Words ?? []).Select(word => new PdfScanWord(word.Text ?? string.Empty, word.X, word.Y, word.Width, word.Height, word.Confidence)).ToArray()))
                .ToArray();
            PdfScanResult result = PdfScanComposer.ComposeAsync(pages, destination!, token).GetAwaiter().GetResult();
            *json = StudioNativeAbi.ToUtf8(JsonSerializer.Serialize(
                new StudioNativeScanResult(StudioNativeAbi.Version, result.Pages, result.SearchablePages, result.Words), StudioNativeJson.Default.StudioNativeScanResult));
        });
    }

    // Documents --------------------------------------------------------------------------------------------------

    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_open")]
    internal static int DocumentOpen(byte* path, byte* password, nint cancellation, nint* document, byte** error) {
        if (document is null) return Fail(error, new ArgumentNullException(nameof(document)));
        *document = 0;
        string? location = StudioNativeAbi.FromUtf8(path);
        string? secret = StudioNativeAbi.FromUtf8(password);
        return Run(error, cancellation, token => {
            StudioNativeDocument opened = StudioNativeDocument.OpenAsync(location!, secret, token).GetAwaiter().GetResult();
            *document = GCHandle.ToIntPtr(GCHandle.Alloc(opened));
        });
    }

    /// <summary>Closes the document and releases its handle. Unsaved changes are discarded; the host asks first.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_close")]
    internal static void DocumentClose(nint document) {
        if (document == 0) return;
        GCHandle handle = GCHandle.FromIntPtr(document);
        (handle.Target as StudioNativeDocument)?.Dispose();
        handle.Free();
    }

    /// <summary>Document facts as UTF-8 JSON: file name and size, revision, unsaved state, undo/redo and page sizes.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_info_json")]
    internal static int DocumentInfoJson(nint document, byte** json, byte** error) {
        if (json is null) return Fail(error, new ArgumentNullException(nameof(json)));
        *json = null;
        return WithDocument(document, error, 0, (target, token) => {
            StudioNativeDocumentInfo info = target.GetInfoAsync(token).GetAwaiter().GetResult();
            *json = StudioNativeAbi.ToUtf8(JsonSerializer.Serialize(info, StudioNativeJson.Default.StudioNativeDocumentInfo));
        });
    }

    /// <summary>Renders a page (numbered from 1) as PNG at <paramref name="scale"/> pixels per point.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_render_page")]
    internal static int DocumentRenderPage(nint document, int page, double scale, nint cancellation,
        byte** png, nuint* length, int* pixelWidth, int* pixelHeight, byte** error) {
        if (png is null || length is null) return Fail(error, new ArgumentNullException(png is null ? nameof(png) : nameof(length)));
        *png = null; *length = 0;
        return WithDocument(document, error, cancellation, (target, token) => {
            PdfRenderedPage rendered = target.RenderPageAsync(page, scale, token).GetAwaiter().GetResult();
            *png = StudioNativeAbi.ToNative(rendered.Bytes);
            *length = (nuint)rendered.Bytes.Length;
            if (pixelWidth is not null) *pixelWidth = rendered.PixelWidth;
            if (pixelHeight is not null) *pixelHeight = rendered.PixelHeight;
        });
    }

    /// <summary>
    /// Renders a rectangle of a page as opaque RGBA pixels for display: <paramref name="x"/>, <paramref name="y"/>,
    /// <paramref name="width"/> by <paramref name="height"/> of the page image at <paramref name="scale"/>.
    /// </summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_render_region")]
    internal static int DocumentRenderRegion(nint document, int page, double scale, int x, int y, int width, int height,
        nint cancellation, byte** pixels, nuint* length, byte** error) {
        if (pixels is null || length is null) return Fail(error, new ArgumentNullException(pixels is null ? nameof(pixels) : nameof(length)));
        *pixels = null; *length = 0;
        return WithDocument(document, error, cancellation, (target, token) => {
            PdfRenderedPixels rendered = target.RenderRegionAsync(page, scale, x, y, width, height, token).GetAwaiter().GetResult();
            *pixels = StudioNativeAbi.ToNative(rendered.Pixels);
            *length = (nuint)rendered.Pixels.Length;
        });
    }

    /// <summary>Rotates pages by a multiple of 90 degrees as one undoable change.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_rotate_pages")]
    internal static int DocumentRotatePages(nint document, int* pages, int count, int degrees, nint cancellation, byte** error) {
        if (pages is null || count <= 0) return Fail(error, new ArgumentException("Choose at least one page."));
        int[] selected = new ReadOnlySpan<int>(pages, count).ToArray();
        return WithDocument(document, error, cancellation, (target, token) => target.RotatePagesAsync(selected, degrees, token).GetAwaiter().GetResult());
    }

    /// <summary>Deletes pages as one undoable change; a document keeps at least one page.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_delete_pages")]
    internal static int DocumentDeletePages(nint document, int* pages, int count, nint cancellation, byte** error) {
        if (pages is null || count <= 0) return Fail(error, new ArgumentException("Choose at least one page."));
        int[] selected = new ReadOnlySpan<int>(pages, count).ToArray();
        return WithDocument(document, error, cancellation, (target, token) => target.DeletePagesAsync(selected, token).GetAwaiter().GetResult());
    }

    /// <summary>Duplicates pages as one undoable change; each copy follows its page.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_duplicate_pages")]
    internal static int DocumentDuplicatePages(nint document, int* pages, int count, nint cancellation, byte** error) {
        if (pages is null || count <= 0) return Fail(error, new ArgumentException("Choose at least one page."));
        int[] selected = new ReadOnlySpan<int>(pages, count).ToArray();
        return WithDocument(document, error, cancellation, (target, token) => target.DuplicatePagesAsync(selected, token).GetAwaiter().GetResult());
    }

    /// <summary>Puts the pages in a new order as one undoable change; <paramref name="order"/> lists every current page once.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_reorder_pages")]
    internal static int DocumentReorderPages(nint document, int* order, int count, nint cancellation, byte** error) {
        if (order is null || count <= 0) return Fail(error, new ArgumentException("The new order must list every page."));
        int[] pages = new ReadOnlySpan<int>(order, count).ToArray();
        return WithDocument(document, error, cancellation, (target, token) => target.ReorderPagesAsync(pages, token).GetAwaiter().GetResult());
    }

    /// <summary>Inserts a blank page of <paramref name="width"/> by <paramref name="height"/> points before a page.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_insert_blank_page")]
    internal static int DocumentInsertBlankPage(nint document, int insertBefore, double width, double height, nint cancellation, byte** error) =>
        WithDocument(document, error, cancellation, (target, token) => target.InsertBlankPageAsync(insertBefore, width, height, token).GetAwaiter().GetResult());

    /// <summary>Writes pages to a new PDF at <paramref name="path"/>, leaving the open document unchanged.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_extract_pages")]
    internal static int DocumentExtractPages(nint document, int* pages, int count, byte* path, nint cancellation, byte** error) {
        if (pages is null || count <= 0) return Fail(error, new ArgumentException("Choose at least one page."));
        int[] selected = new ReadOnlySpan<int>(pages, count).ToArray();
        string? destination = StudioNativeAbi.FromUtf8(path);
        return WithDocument(document, error, cancellation, (target, token) => target.ExtractPagesAsync(selected, destination!, token).GetAwaiter().GetResult());
    }

    /// <summary>
    /// Inserts pages from other PDFs, described by a UTF-8 JSON request, as one undoable change. The result, or for a
    /// source that needs its password the source's position and name, is returned as UTF-8 JSON.
    /// </summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_import_pages_json")]
    internal static int DocumentImportPagesJson(nint document, byte* request, nint cancellation, byte** json, byte** error) {
        if (json is null) return Fail(error, new ArgumentNullException(nameof(json)));
        *json = null;
        StudioNativeImportRequest? parsed;
        try { parsed = JsonSerializer.Deserialize(StudioNativeAbi.FromUtf8(request) ?? "null", StudioNativeJson.Default.StudioNativeImportRequest); }
        catch (JsonException exception) { return Fail(error, new ArgumentException("The import request is not valid JSON.", exception)); }
        if (parsed is null) return Fail(error, new ArgumentNullException(nameof(request)));
        return WithDocument(document, error, cancellation, (target, token) => {
            try {
                StudioNativeImportResult result = target.ImportPagesAsync(parsed, token).GetAwaiter().GetResult();
                *json = StudioNativeAbi.ToUtf8(JsonSerializer.Serialize(result, StudioNativeJson.Default.StudioNativeImportResult));
            } catch (PdfImportPasswordException password) {
                *json = StudioNativeAbi.ToUtf8(JsonSerializer.Serialize(
                    new StudioNativeImportPassword(StudioNativeAbi.Version, password.SourceIndex, password.Name),
                    StudioNativeJson.Default.StudioNativeImportPassword));
                throw;
            }
        });
    }

    /// <summary>
    /// Makes the document searchable as one undoable change: each page is rendered and handed to the host's recognizer,
    /// whose words become invisible text. The recognizer returns words as UTF-8 JSON it allocated with malloc, which the
    /// engine frees; a non-zero status stops the call.
    /// </summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_make_searchable_json")]
    internal static int DocumentMakeSearchableJson(nint document, delegate* unmanaged<nint, int, int, byte*, nuint, byte**, int> recognizer,
        nint context, nint cancellation, byte** json, byte** error) {
        if (json is null) return Fail(error, new ArgumentNullException(nameof(json)));
        *json = null;
        if (recognizer is null) return Fail(error, new ArgumentNullException(nameof(recognizer)));
        return WithDocument(document, error, cancellation, (target, token) => {
            int pageCount = target.PageCount;
            // The OCR runner reports any recognizer failure in general terms, so the host's own failure is kept here.
            Exception? hostFailure = null;
            Func<int, byte[], CancellationToken, IReadOnlyList<PdfScanWord>> recognize = (page, image, pageToken) => {
                pageToken.ThrowIfCancellationRequested();
                byte* words = null;
                int status;
                fixed (byte* png = image) status = recognizer(context, page, pageCount, png, (nuint)image.Length, &words);
                string? text = words is null ? null : StudioNativeAbi.FromUtf8(words);
                if (words is not null) NativeMemory.Free(words);
                if (status == StudioNativeAbi.Cancelled) throw hostFailure = new OperationCanceledException("Text recognition was cancelled.");
                if (status != StudioNativeAbi.Ok) throw hostFailure = new InvalidOperationException($"Text recognition failed on page {page}.");
                try {
                    IReadOnlyList<StudioNativeScanWord>? parsed = JsonSerializer.Deserialize(text ?? "[]",
                        StudioNativeJson.Default.IReadOnlyListStudioNativeScanWord);
                    return parsed?.Select(static word => new PdfScanWord(word.Text ?? string.Empty, word.X, word.Y, word.Width, word.Height, word.Confidence))
                        .ToArray() ?? [];
                } catch (JsonException exception) {
                    throw hostFailure = new ArgumentException($"The recognized words for page {page} are not valid JSON.", exception);
                }
            };
            StudioNativeSearchableResult result;
            try {
                result = target.MakeSearchableAsync(recognize, token).GetAwaiter().GetResult();
            } catch (Exception exception) when (hostFailure is not null && !ReferenceEquals(exception, hostFailure)) {
                throw hostFailure;
            }
            *json = StudioNativeAbi.ToUtf8(JsonSerializer.Serialize(result, StudioNativeJson.Default.StudioNativeSearchableResult));
        });
    }

    /// <summary>The document formats that convert to PDF, as UTF-8 JSON.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_conversion_sources_json")]
    internal static int ConversionSourcesJson(byte** json, byte** error) {
        if (json is null) return Fail(error, new ArgumentNullException(nameof(json)));
        *json = null;
        return Run(error, 0, _ => *json = StudioNativeAbi.ToUtf8(JsonSerializer.Serialize(
            new StudioNativeConversionSources(StudioNativeAbi.Version, PdfConversion.Sources), StudioNativeJson.Default.StudioNativeConversionSources)));
    }

    /// <summary>Converts a document to a PDF at <paramref name="outputPath"/> and returns the route and its notes as UTF-8 JSON.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_convert_to_pdf_json")]
    internal static int ConvertToPdfJson(byte* inputPath, byte* outputPath, nint cancellation, byte** json, byte** error) {
        if (json is null) return Fail(error, new ArgumentNullException(nameof(json)));
        *json = null;
        string input = StudioNativeAbi.FromUtf8(inputPath) ?? string.Empty;
        string output = StudioNativeAbi.FromUtf8(outputPath) ?? string.Empty;
        return Run(error, cancellation, token => {
            PdfConversionResult result = PdfConversion.ConvertAsync(input, output, token).GetAwaiter().GetResult();
            *json = StudioNativeAbi.ToUtf8(JsonSerializer.Serialize(
                new StudioNativeConversionResult(StudioNativeAbi.Version, result.Route, result.Source, result.OutputBytes, result.Notes),
                StudioNativeJson.Default.StudioNativeConversionResult));
        });
    }

    /// <summary>Finds text, ignoring case, and returns the hits with their page rectangles as UTF-8 JSON.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_search_json")]
    internal static int DocumentSearchJson(nint document, byte* query, nint cancellation, byte** json, byte** error) {
        if (json is null) return Fail(error, new ArgumentNullException(nameof(json)));
        *json = null;
        string text = StudioNativeAbi.FromUtf8(query) ?? string.Empty;
        return WithDocument(document, error, cancellation, (target, token) => {
            StudioNativeSearchResult result = target.SearchAsync(text, token).GetAwaiter().GetResult();
            *json = StudioNativeAbi.ToUtf8(JsonSerializer.Serialize(result, StudioNativeJson.Default.StudioNativeSearchResult));
        });
    }

    /// <summary>Adds a markup annotation described by a UTF-8 JSON request as one undoable change.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_add_annotation_json")]
    internal static int DocumentAddAnnotationJson(nint document, byte* request, nint cancellation, byte** error) {
        StudioNativeAnnotationRequest? parsed;
        try { parsed = JsonSerializer.Deserialize(StudioNativeAbi.FromUtf8(request) ?? "null", StudioNativeJson.Default.StudioNativeAnnotationRequest); }
        catch (JsonException exception) { return Fail(error, new ArgumentException("The annotation request is not valid JSON.", exception)); }
        if (parsed is null) return Fail(error, new ArgumentNullException(nameof(request)));
        return WithDocument(document, error, cancellation, (target, token) => target.AddAnnotationAsync(parsed, token).GetAwaiter().GetResult());
    }

    /// <summary>The markup annotations on a page as UTF-8 JSON.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_annotations_json")]
    internal static int DocumentAnnotationsJson(nint document, int page, nint cancellation, byte** json, byte** error) {
        if (json is null) return Fail(error, new ArgumentNullException(nameof(json)));
        *json = null;
        return WithDocument(document, error, cancellation, (target, token) => {
            StudioNativeAnnotationList list = target.GetAnnotationsAsync(page, token).GetAwaiter().GetResult();
            *json = StudioNativeAbi.ToUtf8(JsonSerializer.Serialize(list, StudioNativeJson.Default.StudioNativeAnnotationList));
        });
    }

    /// <summary>Changes an annotation's text, author or colour, described by a UTF-8 JSON update, as one undoable change.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_update_annotation_json")]
    internal static int DocumentUpdateAnnotationJson(nint document, byte* update, nint cancellation, byte** error) {
        StudioNativeAnnotationUpdate? parsed;
        try { parsed = JsonSerializer.Deserialize(StudioNativeAbi.FromUtf8(update) ?? "null", StudioNativeJson.Default.StudioNativeAnnotationUpdate); }
        catch (JsonException exception) { return Fail(error, new ArgumentException("The annotation update is not valid JSON.", exception)); }
        if (parsed is null) return Fail(error, new ArgumentNullException(nameof(update)));
        return WithDocument(document, error, cancellation, (target, token) => target.UpdateAnnotationAsync(parsed, token).GetAwaiter().GetResult());
    }

    /// <summary>Removes an annotation, by its object number from the annotation list, as one undoable change.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_remove_annotation")]
    internal static int DocumentRemoveAnnotation(nint document, int objectNumber, nint cancellation, byte** error) =>
        WithDocument(document, error, cancellation, (target, token) => target.RemoveAnnotationAsync(objectNumber, token).GetAwaiter().GetResult());

    /// <summary>A page's text elements and their boxes as UTF-8 JSON, for selection and copying.</summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_page_text_json")]
    internal static int DocumentPageTextJson(nint document, int page, nint cancellation, byte** json, byte** error) {
        if (json is null) return Fail(error, new ArgumentNullException(nameof(json)));
        *json = null;
        return WithDocument(document, error, cancellation, (target, token) => {
            StudioNativePageText text = target.GetPageTextAsync(page, token).GetAwaiter().GetResult();
            *json = StudioNativeAbi.ToUtf8(JsonSerializer.Serialize(text, StudioNativeJson.Default.StudioNativePageText));
        });
    }

    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_undo")]
    internal static int DocumentUndo(nint document, nint cancellation, byte** error) =>
        WithDocument(document, error, cancellation, (target, token) => target.UndoAsync(token).GetAwaiter().GetResult());

    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_redo")]
    internal static int DocumentRedo(nint document, nint cancellation, byte** error) =>
        WithDocument(document, error, cancellation, (target, token) => target.RedoAsync(token).GetAwaiter().GetResult());

    /// <summary>
    /// Saves to <paramref name="path"/>, or in place when it is null. The file is written beside the destination and
    /// replaces it only when complete, so a failed save leaves the previous file intact.
    /// </summary>
    [UnmanagedCallersOnly(EntryPoint = "officeimo_document_save")]
    internal static int DocumentSave(nint document, byte* path, nint cancellation, byte** error) {
        string? destination = StudioNativeAbi.FromUtf8(path);
        return WithDocument(document, error, cancellation, (target, token) => target.SaveAsync(destination, token).GetAwaiter().GetResult());
    }

    // Plumbing ---------------------------------------------------------------------------------------------------

    private delegate void DocumentCall(StudioNativeDocument document, CancellationToken token);

    private static int WithDocument(nint document, byte** error, nint cancellation, DocumentCall call) {
        if (Resolve<StudioNativeDocument>(document) is not { } target) return Fail(error, new ArgumentException("The document handle is not valid."));
        return Run(error, cancellation, token => call(target, token));
    }

    private static int Run(byte** error, nint cancellation, Action<CancellationToken> call) {
        if (error is not null) *error = null;
        try {
            CancellationToken token = Resolve<CancellationTokenSource>(cancellation)?.Token ?? CancellationToken.None;
            call(token);
            return StudioNativeAbi.Ok;
        } catch (Exception exception) {
            return Fail(error, exception);
        }
    }

    private static int Fail(byte** error, Exception exception) {
        // An exception must never unwind into the host: report it as a status and message.
        if (error is not null) *error = StudioNativeAbi.ToUtf8(exception.Message);
        return StudioNativeAbi.StatusOf(exception);
    }

    private static T? Resolve<T>(nint handle) where T : class {
        if (handle == 0) return null;
        try { return GCHandle.FromIntPtr(handle).Target as T; }
        catch (InvalidOperationException) { return null; }
    }
}

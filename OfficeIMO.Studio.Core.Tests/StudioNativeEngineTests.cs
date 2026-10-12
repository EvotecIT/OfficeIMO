using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Native;

namespace OfficeIMO.Studio.Core.Tests;

/// <summary>The engine library native Apple apps call: the same document journey as the desktop, through its C interface.</summary>
public sealed class StudioNativeEngineTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-native-engine-" + Guid.NewGuid().ToString("N"));

    public StudioNativeEngineTests() {
        Directory.CreateDirectory(_root);
        StudioNativeEngine.Initialize(Path.Combine(_root, "data"));
    }

    public void Dispose() => Directory.Delete(_root, recursive: true);

    private string CreatePdf(int pages = 2) {
        string path = Path.Combine(_root, "source.pdf");
        PdfDocument.Create(document => {
            for (int page = 0; page < pages; page++) document.Page(content => content.Size(300, 400));
        }).Save(path);
        return path;
    }

    [Fact]
    public async Task A_document_opens_renders_rotates_undoes_and_saves_a_copy() {
        string source = CreatePdf();
        byte[] original = File.ReadAllBytes(source);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(source, null, default);

        StudioNativeDocumentInfo info = await document.GetInfoAsync(default);
        Assert.Equal(new[] { 1, 2 }, info.Pages.Select(page => page.Number));
        Assert.False(info.IsDirty);

        var rendered = await document.RenderPageAsync(1, 1.5, default);
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, rendered.Bytes[..4]);
        Assert.Equal(450, rendered.PixelWidth);

        await document.RotatePagesAsync([1], 90, default);
        info = await document.GetInfoAsync(default);
        Assert.Equal(90, info.Pages[0].Rotation);
        Assert.True(info.IsDirty && info.CanUndo);
        await document.UndoAsync(default);
        Assert.Equal(0, (await document.GetInfoAsync(default)).Pages[0].Rotation);
        await document.RedoAsync(default);
        // Rendering follows the edited revision, with the page now landscape.
        Assert.Equal(600, (await document.RenderPageAsync(1, 1.5, default)).PixelWidth);

        string copy = Path.Combine(_root, "rotated.pdf");
        await document.SaveAsync(copy, default);
        info = await document.GetInfoAsync(default);
        Assert.Equal(copy, info.Path);
        Assert.False(info.IsDirty);
        Assert.Equal(original, File.ReadAllBytes(source));
        using StudioNativeDocument reopened = await StudioNativeDocument.OpenAsync(copy, null, default);
        Assert.Equal(90, (await reopened.GetInfoAsync(default)).Pages[0].Rotation);
    }

    // Pages of distinct widths, so a page's width identifies it after an edit.
    private string CreatePagesOfWidths(params double[] widths) {
        string path = Path.Combine(_root, "widths.pdf");
        PdfDocument.Create(document => {
            foreach (double width in widths) document.Page(content => content.Size(width, 400));
        }).Save(path);
        return path;
    }

    private static async Task<double[]> WidthsAsync(StudioNativeDocument document) =>
        (await document.GetInfoAsync(default)).Pages.Select(page => page.Width).ToArray();

    [Fact]
    public async Task Pages_are_deleted_duplicated_reordered_and_inserted_as_undoable_changes() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePagesOfWidths(300, 310, 320), null, default);
        Assert.True((await document.GetInfoAsync(default)).CanMutatePages);

        await document.DeletePagesAsync([2], default);
        Assert.Equal(new[] { 300D, 320D }, await WidthsAsync(document));
        await document.UndoAsync(default);
        Assert.Equal(new[] { 300D, 310D, 320D }, await WidthsAsync(document));

        await document.DuplicatePagesAsync([1], default);
        Assert.Equal(new[] { 300D, 300D, 310D, 320D }, await WidthsAsync(document));

        await document.ReorderPagesAsync([4, 3, 1, 2], default);
        Assert.Equal(new[] { 320D, 310D, 300D, 300D }, await WidthsAsync(document));

        await document.InsertBlankPageAsync(5, 500, 600, default);
        StudioNativeDocumentInfo info = await document.GetInfoAsync(default);
        Assert.Equal((500D, 600D), (info.Pages[^1].Width, info.Pages[^1].Height));

        // Each edit undoes on its own, back to the original pages.
        for (int edit = 0; edit < 3; edit++) await document.UndoAsync(default);
        Assert.Equal(new[] { 300D, 310D, 320D }, await WidthsAsync(document));
    }

    [Fact]
    public async Task Page_edits_refuse_what_would_damage_the_document() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePagesOfWidths(300, 310), null, default);

        await Assert.ThrowsAsync<ArgumentException>(() => document.DeletePagesAsync([1, 2], default));
        await Assert.ThrowsAsync<ArgumentException>(() => document.ReorderPagesAsync([1, 1], default));
        await Assert.ThrowsAsync<ArgumentException>(() => document.ReorderPagesAsync([2], default));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => document.DuplicatePagesAsync([3], default));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => document.InsertBlankPageAsync(4, 300, 400, default));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => document.InsertBlankPageAsync(1, 0, 400, default));

        Assert.False((await document.GetInfoAsync(default)).IsDirty);
    }

    [Fact]
    public async Task Extracted_pages_form_a_new_document_and_leave_the_open_one_unchanged() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePagesOfWidths(300, 310, 320), null, default);
        long revision = (await document.GetInfoAsync(default)).Revision;
        string extracted = Path.Combine(_root, "extracted.pdf");

        await document.ExtractPagesAsync([3, 1], extracted, default);

        using StudioNativeDocument copy = await StudioNativeDocument.OpenAsync(extracted, null, default);
        Assert.Equal(new[] { 300D, 320D }, await WidthsAsync(copy));
        StudioNativeDocumentInfo info = await document.GetInfoAsync(default);
        Assert.Equal(revision, info.Revision);
        Assert.False(info.IsDirty);
    }

    [Fact]
    public async Task Search_finds_text_with_rectangles_on_the_page_as_displayed() {
        string path = Path.Combine(_root, "text.pdf");
        PdfDocument.Create(document => {
            foreach (string text in new[] { "Alpha needle here", "Nothing to see", "Another NEEDLE" })
                document.Page(page => page.Size(400, 300).Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(text)))));
        }).Save(path);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);

        StudioNativeSearchResult result = await document.SearchAsync("  needle ", default);

        Assert.Equal("needle", result.Query);
        Assert.Equal(new[] { (1, 1), (3, 2) }, result.Hits.Select(hit => (hit.Page, hit.Occurrence)));
        Assert.All(result.Hits, hit => {
            Assert.True(hit.Bounds.Width > 0 && hit.Bounds.Height > 0);
            Assert.InRange(hit.Bounds.X + hit.Bounds.Width, 0D, 400D);
            Assert.InRange(hit.Bounds.Y + hit.Bounds.Height, 0D, 300D);
            Assert.Single(hit.LineBounds);
        });

        // Rotating the page turns the rectangles with it, into the displayed page's space.
        await document.RotatePagesAsync([1], 90, default);
        StudioNativeSearchHit rotated = (await document.SearchAsync("needle", default)).Hits[0];
        Assert.True(rotated.Bounds.Height > rotated.Bounds.Width);
        Assert.InRange(rotated.Bounds.Y + rotated.Bounds.Height, 0D, 400D);
    }

    [Fact]
    public unsafe void The_C_interface_reorders_pages_and_returns_search_hits_as_json() {
        string path = Path.Combine(_root, "text.pdf");
        PdfDocument.Create(document => {
            foreach (string text in new[] { "first page", "second needle" })
                document.Page(page => page.Size(400, 300).Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(text)))));
        }).Save(path);
        delegate* unmanaged<byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpen;
        delegate* unmanaged<nint, int*, int, nint, byte**, int> reorder = &StudioNativeExports.DocumentReorderPages;
        delegate* unmanaged<nint, byte*, nint, byte**, byte**, int> search = &StudioNativeExports.DocumentSearchJson;
        delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;

        nint document = 0; byte* error = null;
        fixed (byte* location = Encoding.UTF8.GetBytes(path + "\0")) Assert.Equal(StudioNativeAbi.Ok, open(location, null, 0, &document, &error));
        try {
            int* order = stackalloc int[] { 2, 1 };
            Assert.Equal(StudioNativeAbi.Ok, reorder(document, order, 2, 0, &error));
            Assert.Equal(StudioNativeAbi.InvalidArgument, reorder(document, order, 1, 0, &error));
            free(error);

            byte* json = null;
            fixed (byte* query = Encoding.UTF8.GetBytes("needle\0")) Assert.Equal(StudioNativeAbi.Ok, search(document, query, 0, &json, &error));
            using JsonDocument hits = JsonDocument.Parse(Marshal.PtrToStringUTF8((nint)json)!);
            free(json);
            JsonElement hit = Assert.Single(hits.RootElement.GetProperty("hits").EnumerateArray());
            Assert.Equal(1, hit.GetProperty("page").GetInt32());
            Assert.True(hit.GetProperty("bounds").GetProperty("width").GetDouble() > 0);
            Assert.Equal(1, hits.RootElement.GetProperty("abiVersion").GetInt32());
        } finally { close(document); }
    }

    [Fact]
    public async Task Page_sizes_are_the_rendered_size_at_the_page_user_unit() {
        string path = Path.Combine(_root, "user-unit.pdf");
        File.WriteAllBytes(path, Encoding.ASCII.GetBytes("""
            %PDF-1.7
            1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
            2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj
            3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /CropBox [0 0 200 300] /UserUnit 2 /Resources << >> >> endobj
            trailer << /Root 1 0 R /Size 4 >>
            %%EOF
            """));
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);

        StudioNativePageInfo page = (await document.GetInfoAsync(default)).Pages[0];

        Assert.Equal((400D, 600D), (page.Width, page.Height));
        Assert.Equal(400, (await document.RenderPageAsync(1, 1D, default)).PixelWidth);
    }

    [Fact]
    public async Task Closing_discards_the_recovery_copy_of_unsaved_edits() {
        string source = CreatePdf();
        StudioNativeDocument document = await StudioNativeDocument.OpenAsync(source, null, default);
        await document.RotatePagesAsync([1], 90, default);
        Assert.NotEmpty(Directory.EnumerateFiles(Path.Combine(_root, "data", "Recovery")));

        document.Dispose();

        Assert.Empty(Directory.EnumerateFiles(Path.Combine(_root, "data", "Recovery"), "*.pdf"));
        using StudioNativeDocument reopened = await StudioNativeDocument.OpenAsync(source, null, default);
        Assert.False((await reopened.GetInfoAsync(default)).HasRecovery);
    }

    [Fact]
    public async Task Tiles_beyond_the_page_budget_join_into_the_zoomed_page() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePdf(1), null, default);
        // 300 x 400 points at 10x is 12 million pixels, beyond the 8.4 million pixel budget, so it renders as tiles.
        const double scale = 10D;
        var tiles = new List<PdfRenderedPixels>();
        foreach ((int x, int y) in new[] { (0, 0), (1500, 0), (0, 2000), (1500, 2000) })
            tiles.Add(await document.RenderRegionAsync(1, scale, x, y, 1500, 2000, default));

        Assert.All(tiles, tile => Assert.Equal(1500 * 2000 * 4, tile.Pixels.Length));
        Assert.All(tiles, tile => Assert.True(tile.Pixels.Where((_, index) => index % 4 == 3).All(alpha => alpha == 255)));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => document.RenderRegionAsync(1, scale, 0, 0, 3000, 4000, default));
    }

    [Fact]
    public unsafe void The_C_interface_renders_page_rectangles_as_pixels() {
        string source = CreatePdf(1);
        delegate* unmanaged<byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpen;
        delegate* unmanaged<nint, int, double, int, int, int, int, nint, byte**, nuint*, byte**, int> region = &StudioNativeExports.DocumentRenderRegion;
        delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;

        nint document = 0; byte* error = null;
        fixed (byte* path = Encoding.UTF8.GetBytes(source + "\0")) Assert.Equal(StudioNativeAbi.Ok, open(path, null, 0, &document, &error));
        try {
            byte* pixels = null; nuint length = 0;
            Assert.Equal(StudioNativeAbi.Ok, region(document, 1, 1.5D, 0, 0, 450, 600, 0, &pixels, &length, &error));
            Assert.Equal((nuint)(450 * 600 * 4), length);
            Assert.Equal(255, pixels[3]);
            free(pixels);

            // Outside the page image, or more than the budget in one rectangle, is the host's argument to correct.
            Assert.Equal(StudioNativeAbi.InvalidArgument, region(document, 1, 1.5D, 400, 0, 100, 100, 0, &pixels, &length, &error));
            free(error);
            Assert.Equal(StudioNativeAbi.InvalidArgument, region(document, 1, 20D, 0, 0, 6000, 8000, 0, &pixels, &length, &error));
            free(error);
            Assert.True(pixels is null);
        } finally { close(document); }
    }

    [Fact]
    public unsafe void The_C_interface_reports_failures_as_status_and_message_without_throwing() {
        string source = CreatePdf(1);
        delegate* unmanaged<byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpen;
        delegate* unmanaged<nint, int, double, nint, byte**, nuint*, int*, int*, byte**, int> render = &StudioNativeExports.DocumentRenderPage;
        delegate* unmanaged<nint, byte**, byte**, int> info = &StudioNativeExports.DocumentInfoJson;
        delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;
        delegate* unmanaged<nint> createCancellation = &StudioNativeExports.CancellationCreate;
        delegate* unmanaged<nint, void> cancel = &StudioNativeExports.CancellationCancel;
        delegate* unmanaged<nint, void> freeCancellation = &StudioNativeExports.CancellationFree;

        string Take(byte* text) { string value = Marshal.PtrToStringUTF8((nint)text)!; free(text); return value; }
        byte[] Utf8(string value) => Encoding.UTF8.GetBytes(value + "\0");

        nint document = 0; byte* error = null;
        byte[] missing = Utf8(Path.Combine(_root, "missing.pdf"));
        fixed (byte* path = missing) Assert.Equal(StudioNativeAbi.NotFound, open(path, null, 0, &document, &error));
        Assert.Equal(0, document);
        Assert.NotEmpty(Take(error));

        fixed (byte* path = Utf8(source)) Assert.Equal(StudioNativeAbi.Ok, open(path, null, 0, &document, &error));
        Assert.True(error is null);
        try {
            byte* json = null;
            Assert.Equal(StudioNativeAbi.Ok, info(document, &json, &error));
            using (JsonDocument facts = JsonDocument.Parse(Take(json))) {
                Assert.Equal(StudioNativeAbi.Version, facts.RootElement.GetProperty("abiVersion").GetInt32());
                Assert.Equal(1, facts.RootElement.GetProperty("pages").GetArrayLength());
                Assert.Equal(StudioPdfSecurityPolicy.MaximumRasterPixels, facts.RootElement.GetProperty("maxRenderPixels").GetInt64());
            }

            byte* png = null; nuint length = 0; int width = 0, height = 0;
            Assert.Equal(StudioNativeAbi.Ok, render(document, 1, 1D, 0, &png, &length, &width, &height, &error));
            Assert.Equal((nuint)new ReadOnlySpan<byte>(png, (int)length).Length, length);
            Assert.Equal((300, 400), (width, height));
            free(png);

            Assert.Equal(StudioNativeAbi.InvalidArgument, render(document, 5, 1D, 0, &png, &length, null, null, &error));
            Assert.True(png is null);
            Take(error);

            // A page image beyond the pixel budget is the host's argument to correct, not an engine failure.
            Assert.Equal(StudioNativeAbi.InvalidArgument, render(document, 1, 20D, 0, &png, &length, null, null, &error));
            Assert.True(png is null);
            Take(error);

            nint cancellation = createCancellation();
            cancel(cancellation);
            Assert.Equal(StudioNativeAbi.Cancelled, render(document, 1, 1D, cancellation, &png, &length, null, null, null));
            freeCancellation(cancellation);

            // A stale or invented handle is refused rather than dereferenced.
            Assert.Equal(StudioNativeAbi.InvalidArgument, info(0, &json, &error));
            Take(error);
        } finally { close(document); }
    }

    [Theory]
    [InlineData(typeof(OperationCanceledException), StudioNativeAbi.Cancelled)]
    [InlineData(typeof(FileNotFoundException), StudioNativeAbi.NotFound)]
    [InlineData(typeof(ArgumentOutOfRangeException), StudioNativeAbi.InvalidArgument)]
    [InlineData(typeof(NotSupportedException), StudioNativeAbi.Unsupported)]
    [InlineData(typeof(StudioNativeNotInitializedException), StudioNativeAbi.NotInitialized)]
    [InlineData(typeof(IOException), StudioNativeAbi.Failed)]
    public void Engine_errors_map_to_stable_status_codes(Type exception, int status) =>
        Assert.Equal(status, StudioNativeAbi.StatusOf((Exception)Activator.CreateInstance(exception)!));

    [Fact]
    public void Password_errors_ask_the_host_to_prompt() {
        Assert.Equal(StudioNativeAbi.PasswordRequired, StudioNativeAbi.StatusOf(new PdfPasswordRequiredException("Password required.")));
        Assert.Equal(StudioNativeAbi.InvalidPassword, StudioNativeAbi.StatusOf(new PdfInvalidPasswordException("Wrong password.")));
        Assert.Equal(StudioNativeAbi.Unsupported, StudioNativeAbi.StatusOf(new PdfUnsupportedEncryptionException("Unsupported encryption.")));
    }
}

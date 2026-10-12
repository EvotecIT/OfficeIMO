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

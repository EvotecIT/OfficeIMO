using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
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
    public async Task Rotation_preserves_stream_and_tagged_sources_through_native_workspace_history() {
        foreach (string name in new[] { "latin-0-native.pdf", "openpreserve-pdfa1b-text.pdf" }) {
            string source = Path.Combine(_root, name);
            File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", name), source);
            byte[] original = File.ReadAllBytes(source);
            using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(source, null, default);
            await document.RotatePagesAsync([1], 90, default);
            Assert.Equal(90, (await document.GetInfoAsync(default)).Pages[0].Rotation);
            await document.RotatePagesAsync([1], 90, default);
            Assert.Equal(180, (await document.GetInfoAsync(default)).Pages[0].Rotation);
            await document.UndoAsync(default);
            Assert.Equal(90, (await document.GetInfoAsync(default)).Pages[0].Rotation);
            await document.UndoAsync(default);
            Assert.Equal(0, (await document.GetInfoAsync(default)).Pages[0].Rotation);
            await document.RedoAsync(default);
            string copy = Path.Combine(_root, "rotated-" + name);
            await document.SaveAsync(copy, default);
            byte[] output = File.ReadAllBytes(copy);
            Assert.True(output.AsSpan(0, original.Length).SequenceEqual(original));
            using StudioNativeDocument reopened = await StudioNativeDocument.OpenAsync(copy, null, default);
            Assert.Equal(90, (await reopened.GetInfoAsync(default)).Pages[0].Rotation);
            Assert.Equal(original, File.ReadAllBytes(source));
        }
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
            }

            byte* png = null; nuint length = 0; int width = 0, height = 0;
            Assert.Equal(StudioNativeAbi.Ok, render(document, 1, 1D, 0, &png, &length, &width, &height, &error));
            Assert.Equal((nuint)new ReadOnlySpan<byte>(png, (int)length).Length, length);
            Assert.Equal((300, 400), (width, height));
            free(png);

            Assert.Equal(StudioNativeAbi.InvalidArgument, render(document, 5, 1D, 0, &png, &length, null, null, &error));
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

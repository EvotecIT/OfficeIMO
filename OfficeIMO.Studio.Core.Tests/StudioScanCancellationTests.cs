using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Native;

namespace OfficeIMO.Studio.Core.Tests;

public sealed class StudioScanCancellationTests {
    [Fact]
    public async Task Cancelled_native_scan_preserves_existing_output_and_a_fresh_token_can_compose() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-scan-cancellation-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            StudioNativeEngine.Initialize(Path.Combine(root, "data"));
            string source = Path.Combine(root, "source.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(300, 400))).Save(source);
            string image = Path.Combine(root, "page.png");
            using (StudioNativeDocument document = await StudioNativeDocument.OpenAsync(source, null, default)) {
                File.WriteAllBytes(image, (await document.RenderPageAsync(1, 1D, default)).Bytes);
            }
            string output = Path.Combine(root, "scan.pdf");
            byte[] original = Encoding.UTF8.GetBytes("existing publication");
            File.WriteAllBytes(output, original);
            string request = JsonSerializer.Serialize(new { pages = new[] { new { image, words = Array.Empty<object>() } } });

            Invoke(request, output, cancelled: true);
            Assert.Equal(original, File.ReadAllBytes(output));
            string fresh = Path.Combine(root, "fresh.pdf");
            Invoke(request, fresh, cancelled: false);
            using StudioNativeDocument scan = await StudioNativeDocument.OpenAsync(fresh, null, default);
            Assert.Single((await scan.GetInfoAsync(default)).Pages);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private static unsafe void Invoke(string request, string output, bool cancelled) {
        delegate* unmanaged<nint> create = &StudioNativeExports.CancellationCreate;
        delegate* unmanaged<nint, void> cancel = &StudioNativeExports.CancellationCancel;
        delegate* unmanaged<nint, void> freeToken = &StudioNativeExports.CancellationFree;
        delegate* unmanaged<byte*, byte*, nint, byte**, byte**, int> compose = &StudioNativeExports.CreatePdfFromScanJson;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;
        nint token = create();
        byte* json = null; byte* error = null;
        try {
            if (cancelled) cancel(token);
            fixed (byte* input = Encoding.UTF8.GetBytes(request + "\0"))
            fixed (byte* destination = Encoding.UTF8.GetBytes(output + "\0")) {
                int status = compose(input, destination, token, &json, &error);
                Assert.Equal(cancelled ? StudioNativeAbi.Cancelled : StudioNativeAbi.Ok, status);
            }
            if (cancelled) Assert.True(json is null);
            else {
                using JsonDocument summary = JsonDocument.Parse(Marshal.PtrToStringUTF8((nint)json)!);
                Assert.Equal(1, summary.RootElement.GetProperty("pages").GetInt32());
            }
        } finally {
            free(json); free(error); freeToken(token);
        }
    }
}

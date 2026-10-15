using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Native;

namespace OfficeIMO.Studio.Core.Tests;

public sealed class StudioNativeRedactionWireTests {
    [Fact]
    public unsafe void C_redaction_requires_current_revision_after_another_thread_edits_the_handle() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-redaction-wire-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        StudioNativeEngine.Initialize(Path.Combine(root, "data"));
        string source = Path.Combine(root, "source.pdf");
        PdfDocument.Create(document => {
            foreach (string text in new[] { "First secret", "Second secret" })
                document.Page(page => page.Size(400, 300).Content(content => content.Paragraph(p => p.Text(text))));
        }).Save(source);
        delegate* unmanaged<byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpen;
        delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;
        nint handle = 0;
        byte* path = StudioNativeAbi.ToUtf8(source);
        byte* error = null;
        try {
            Assert.Equal(StudioNativeAbi.Ok, open(path, null, 0, &handle, &error));
            var first = Request(handle, "{\"find\":{\"text\":\"secret\"}}");
            Assert.Equal(StudioNativeAbi.Ok, first.Status);
            using var found = JsonDocument.Parse(first.Json!);
            Assert.Equal(StudioNativeAbi.Version, found.RootElement.GetProperty("abiVersion").GetInt32());
            long revision = found.RootElement.GetProperty("revision").GetInt64();
            Assert.Equal(2, found.RootElement.GetProperty("marks").GetArrayLength());
            var missing = Request(handle, "{\"apply\":{\"find\":{\"text\":\"secret\"},\"include\":[0]}}");
            Assert.Equal(StudioNativeAbi.InvalidArgument, missing.Status);
            Assert.Null(missing.Json);
            Assert.Contains("revision", missing.Error);

            int changed = -1;
            string? changeError = null;
            nint sharedHandle = handle;
            var other = new Thread(() => {
                delegate* unmanaged<nint, int*, int, int, nint, byte**, int> rotate = &StudioNativeExports.DocumentRotatePages;
                delegate* unmanaged<void*, void> release = &StudioNativeExports.Free;
                int page = 1;
                byte* failure = null;
                changed = rotate(sharedHandle, &page, 1, 90, 0, &failure);
                changeError = StudioNativeAbi.FromUtf8(failure);
                release(failure);
            });
            other.Start();
            other.Join();
            Assert.Equal(StudioNativeAbi.Ok, changed);
            Assert.Null(changeError);
            var stale = Request(handle, JsonSerializer.Serialize(new { apply = new { revision, find = new { text = "secret" }, include = new[] { 0 } } }));
            Assert.Equal(StudioNativeAbi.Failed, stale.Status);
            Assert.Null(stale.Json);
            Assert.Contains("changed", stale.Error);
            var current = Request(handle, "{\"find\":{\"text\":\"secret\"}}");
            using var refreshed = JsonDocument.Parse(current.Json!);
            long next = refreshed.RootElement.GetProperty("revision").GetInt64();
            Assert.NotEqual(revision, next);
            Assert.Equal(2, refreshed.RootElement.GetProperty("marks").GetArrayLength());
            var applied = Request(handle, JsonSerializer.Serialize(new { apply = new { revision = next, find = new { text = "secret" }, include = new[] { 0 } } }));
            Assert.Equal(StudioNativeAbi.Ok, applied.Status);
            using var verified = JsonDocument.Parse(applied.Json!);
            Assert.True(verified.RootElement.GetProperty("verification").GetProperty("verified").GetBoolean());
            var remaining = Request(handle, "{\"find\":{\"text\":\"secret\"}}");
            using var matches = JsonDocument.Parse(remaining.Json!);
            Assert.Equal(1, matches.RootElement.GetProperty("marks").GetArrayLength());
            Assert.Equal(2, matches.RootElement.GetProperty("marks")[0].GetProperty("page").GetInt32());
        } finally {
            free(path);
            free(error);
            if (handle != 0) close(handle);
            Directory.Delete(root, true);
        }
    }

    private static unsafe (int Status, string? Json, string? Error) Request(nint handle, string request) {
        delegate* unmanaged<nint, byte*, nint, byte**, byte**, int> call = &StudioNativeExports.DocumentRedactJson;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;
        byte* input = StudioNativeAbi.ToUtf8(request);
        byte* json = null;
        byte* error = null;
        try {
            int status = call(handle, input, 0, &json, &error);
            return (status, StudioNativeAbi.FromUtf8(json), StudioNativeAbi.FromUtf8(error));
        } finally { free(input); free(json); free(error); }
    }
}

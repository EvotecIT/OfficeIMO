using System.Reflection;
using System.Text;
using System.Text.Json;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Platform.Storage;
using OfficeIMO.Studio.Features.Mobile;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioMobileOutputTests {
    [Fact]
    public async Task CancellingDuringProviderImportLeavesNoWorkingCopy() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-import-cancel-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "documents")));
                var input = new TestStorageFile("content://mobile/source.html", Encoding.UTF8.GetBytes("<html>source</html>"), "source.html");
                var provider = DispatchProxy.Create<IStorageProvider, TestStorageFile.StorageProxy>();
                ((TestStorageFile.StorageProxy)(object)provider).Call = (method, _) => method switch {
                    "OpenFilePickerAsync" => Task.FromResult<IReadOnlyList<IStorageFile>>([input.Item]),
                    _ => throw new NotSupportedException(method)
                };
                var view = new MobileWorkspaceView();
                var host = new MobileDocumentHost(services, view, _ => Task.CompletedTask, () => provider);
                using var controller = new MobileDocumentController(services, _ => Task.FromResult<IStorageFile?>(null), _ => Task.CompletedTask, host);
                var model = controller.Document.ProvenanceWorkbench;
                input.BeforeRead = () => model.CancelCommand.Execute(null);
                await model.ChooseInputCommand.ExecuteAsync(null);
                Assert.False(model.IsBusy);
                Assert.Empty(model.InputPath);
                Assert.Contains("cancelled", model.Status, StringComparison.OrdinalIgnoreCase);
                Assert.False(Directory.Exists(services.LocalDocuments!.Path));
                Assert.Equal(input.Reads, input.ClosedReads);
                Assert.Equal(0, input.Writes);
                services.Storage.Dispose();
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(1366, 1024)]
    [InlineData(390, 844)]
    public async Task ProvenanceImportsACopyAndSharesItsReportAndProviderOutputs(int width, int height) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-provenance-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "documents")));
                byte[] original = Encoding.UTF8.GetBytes("<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>review\u202Ethis</body></html>");
                var input = new TestStorageFile("content://mobile/source.html", original, "source.html");
                var provider = DispatchProxy.Create<IStorageProvider, TestStorageFile.StorageProxy>();
                ((TestStorageFile.StorageProxy)(object)provider).Call = (method, _) => method switch {
                    "OpenFilePickerAsync" => Task.FromResult<IReadOnlyList<IStorageFile>>([input.Item]),
                    _ => throw new NotSupportedException(method)
                };
                string? shared = null;
                byte[]? sharedBytes = null;
                Task Share(string path) { shared = path; sharedBytes = File.ReadAllBytes(path); return Task.CompletedTask; }
                var view = new MobileWorkspaceView();
                var host = new MobileDocumentHost(services, view, Share, () => provider);
                using var controller = new MobileDocumentController(services, _ => Task.FromResult<IStorageFile?>(null), Share, host);
                view.Connect(controller);
                var window = new Window { Content = view, Width = width, Height = height };
                try {
                    window.Show();
                    var document = controller.Document;
                    document.ShowProvenanceCommand.Execute(null);
                    var model = document.ProvenanceWorkbench;
                    await model.ChooseInputCommand.ExecuteAsync(null);
                    Assert.True(model.UsesWorkingCopies);
                    Assert.StartsWith(services.LocalDocuments!.Path, model.InputPath);
                    Assert.Equal(original, File.ReadAllBytes(model.InputPath));
                    await model.AssessCommand.ExecuteAsync(null);
                    Assert.Contains(model.Findings, finding => finding.Contains("U+202E"));
                    model.RemoveReferences = true;
                    await model.CreateCopyCommand.ExecuteAsync(null);
                    Assert.True(File.Exists(model.OutputPath), model.Status);
                    Assert.DoesNotContain("c2pa-manifest", File.ReadAllText(model.OutputPath));
                    Assert.Equal(original, input.Bytes);
                    Assert.Equal(original, File.ReadAllBytes(model.InputPath));
                    await model.ExportReportCommand.ExecuteAsync(null);
                    var job = services.Jobs.Entries[0];
                    Assert.Equal(model.ReportPath, job.OutputPath);
                    await document.Jobs.OpenOutputCommand.ExecuteAsync(job);
                    Assert.Null(document.Jobs.ActionError);
                    Assert.Equal(model.ReportPath, shared);
                    using var json = JsonDocument.Parse(sharedBytes!);
                    Assert.Equal(model.OutputHash, json.RootElement.GetProperty("outputSha256").GetString());
                    window.UpdateLayout();
                    AvaloniaHeadlessPlatform.ForceRenderTimerTick();
                    using var frame = window.CaptureRenderedFrame();
                    Assert.NotNull(frame);
                    if (Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT") is { } evidence) {
                        Directory.CreateDirectory(evidence);
                        frame.Save(Path.Combine(evidence, "application-provenance-" + width + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                    }
                    // Non-file provider identities also share a bounded local snapshot, not a URI launch.
                    await host.OpenUriAsync(input.Location);
                    Assert.Equal(original, sharedBytes);
                    Assert.False(File.Exists(shared));
                    var folder = new StudioProviderOutputFolderTests.OutputFolder();
                    string location = (await services.Storage.RegisterFolderAsync([folder.Item], default))!;
                    var error = await Assert.ThrowsAsync<IOException>(() => host.OpenUriAsync(new Uri(location)));
                    Assert.Contains("Files", error.Message);
                } finally { window.Close(); services.Storage.Dispose(); }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }
}

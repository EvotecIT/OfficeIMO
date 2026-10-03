using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Infrastructure;
using Avalonia.Platform.Storage;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioIWorkProviderPackageTests {
    [Theory]
    [InlineData(false, 840)]
    [InlineData(false, 1280)]
    [InlineData(true, 840)]
    public async Task Registered_package_uses_folder_capture_and_rejects_replacement(bool replace, int width) {
        // On other hosts local folders use the existing local-package route; macOS selects provider access.
        if (!OperatingSystem.IsMacOS()) return;
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string root = Path.Combine(services.Paths.Root, "selected.numbers");
            Directory.CreateDirectory(root);
            ZipFile.ExtractToDirectory(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IWork", "simple.numbers"), root);
            var folder = CreateFolder(root);
            await services.Storage.RegisterManyAsync([folder], default);
            var guard = new MutationGuard(() => {
                if (!replace) return;
                string original = root + ".old";
                Directory.Move(root, original);
                Directory.CreateDirectory(root);
                foreach (var file in Directory.GetFiles(original, "*", SearchOption.AllDirectories)) {
                    string destination = Path.Combine(root, Path.GetRelativePath(original, file));
                    Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
                    File.Copy(file, destination);
                }
            });
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services, publicationGuard: guard);
            model.FileDialogs = new PackageDialogs(root);
            var queue = model.ConversionWorkbench;
            queue.OutputFolder = services.Paths.Root;
            var view = new ConversionWorkbenchView { DataContext = model };
            var window = new Window { Width = width, Height = 720, Content = view };
            try {
                window.Show(); window.UpdateLayout();
                var button = Assert.Single(view.GetVisualDescendants().OfType<Button>(), b => ReferenceEquals(b.Command, queue.AddPackageCommand));
                Assert.True(button.IsEnabled);
                Point point = button.TranslatePoint(new Point(button.Bounds.Width / 2, button.Bounds.Height / 2), window)!.Value;
                Assert.InRange(point.X, 0, window.Bounds.Width);
                window.MouseDown(point, MouseButton.Left);
                window.MouseUp(point, MouseButton.Left);
                Assert.Single(queue.Jobs);
                window.UpdateLayout();
                string? visual = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrWhiteSpace(visual)) {
                    Directory.CreateDirectory(visual);
                    using var frame = window.CaptureRenderedFrame();
                    Assert.NotNull(frame);
                    frame.Save(Path.Combine(visual, $"package-intake-{width}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
            } finally { window.Close(); }
            var job = Assert.Single(queue.Jobs);
            job.AllowPartialEditableReconstruction = true;
            await queue.RunQueueCommand.ExecuteAsync(null);
            Assert.Equal(replace ? ConversionJobState.Failed : ConversionJobState.Completed, job.State);
            Assert.Equal(!replace, job.HasOutput);
            Assert.Contains(job.Diagnostics, d => d.Code == "SourceSnapshot" && d.Details["snapshotKind"] == "DirectoryPackage");
            if (replace) Assert.Contains("replaced", job.Summary!, StringComparison.OrdinalIgnoreCase);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task Package_root_guard_rejects_nested_output_and_releases_iterator() {
        if (!OperatingSystem.IsMacOS()) return;
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string root = Path.Combine(services.Paths.Root, "selected.pages");
            Directory.CreateDirectory(root);
            File.WriteAllText(Path.Combine(root, "member.txt"), "keep");
            int active = 0;
            var folder = CreateFolder(root, delta => active += delta);
            await services.Storage.RegisterFolderAsync([folder], default);
            using var inputs = services.Storage.CreateDirectoryInputs([root]);
            var package = Assert.IsType<OfficeWorkflowDirectoryPackageInput>(inputs.CreatePackageInput(root));
            Assert.False(await package.SourcePublicationGuard.CanPublishAsync(Path.Combine(root, "new.docx"), false, default));
            Assert.Equal(0, active);
            Assert.True(await package.SourcePublicationGuard.CanPublishAsync(Path.Combine(services.Paths.Root, "new.docx"), false, default));
            Assert.Equal(0, active);
            string alias = Path.Combine(services.Paths.Root, "alias.docx");
            Assert.Equal(0, Link(Path.Combine(root, "member.txt"), alias));
            Assert.False(await package.SourcePublicationGuard.CanPublishAsync(alias, false, default));
            Assert.Equal(0, active);
            return true;
        }, CancellationToken.None);
    }

    [DllImport("libc", EntryPoint = "link", SetLastError = true)]
    private static extern int Link(string source, string destination);

    private sealed class PackageDialogs(string path) : IStudioFileDialogs {
        public Task<string?> PickFolderAsync(string title, CancellationToken token) => Task.FromResult<string?>(path);
        public Task<string?> PickOpenFileAsync(string title, StudioFileType type, CancellationToken token) => throw new NotSupportedException();
        public Task<string?> PickSaveFileAsync(string title, string name, StudioFileType type, CancellationToken token) => throw new NotSupportedException();
    }

    private sealed class MutationGuard(Action action) : IOfficeWorkflowPublicationGuard {
        private bool _done;
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            if (!_done) { _done = true; action(); }
            return ValueTask.FromResult(true);
        }
    }

    private static IStorageFolder CreateFolder(string path, Action<int>? scope = null) {
        var item = DispatchProxy.Create<IStorageFolder, TestStorageFile.StorageProxy>();
        ((TestStorageFile.StorageProxy)(object)item).Call = (method, _) => method switch {
            "get_Name" => Path.GetFileName(path), "get_Path" => new Uri(path), "get_CanBookmark" => false,
            "GetItemsAsync" => Enumerate(), "Dispose" => null, _ => throw new NotSupportedException(method)
        };
        return item;
        async IAsyncEnumerable<IStorageItem> Enumerate() {
            scope?.Invoke(1);
            try {
                foreach (string directory in Directory.GetDirectories(path)) yield return CreateFolder(directory, scope);
                foreach (string file in Directory.GetFiles(path)) {
                    var child = DispatchProxy.Create<IStorageFile, TestStorageFile.StorageProxy>();
                    ((TestStorageFile.StorageProxy)(object)child).Call = (method, _) => method switch {
                        "get_Name" => Path.GetFileName(file), "get_Path" => new Uri(file), "get_CanBookmark" => false,
                        "OpenReadAsync" => Task.FromResult<Stream>(File.OpenRead(file)), "Dispose" => null,
                        _ => throw new NotSupportedException(method)
                    };
                    yield return child;
                }
                await Task.CompletedTask;
            } finally { scope?.Invoke(-1); }
        }
    }
}

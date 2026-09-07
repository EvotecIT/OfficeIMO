using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Platform.Storage;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioProviderFolderTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 820, true)]
    public async Task AssemblyUsesProviderFolderAndDisposesMembersAfterPublication(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            var first = new TestStorageFile("content://folder/first", StudioProviderDocumentTests.CreatePdf(), "First.pdf");
            var second = new TestStorageFile("content://folder/second", StudioProviderDocumentTests.CreatePdf(2), "Second.pdf");
            var nested = new Folder("content://folder/nested", "Nested documents", [second.Item]);
            var root = new Folder("content://folder/root", "Selected provider folder with nested documents", [first.Item, nested.Item]);
            string? location = await services.Storage.RegisterFolderAsync([root.Item], default);
            string output = Path.Combine(services.Paths.Root, "provider-folder-assembly.pdf");
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickAssemblyFolder: _ => Task.FromResult(location), pickSavePdf: _ => Task.FromResult<string?>(output), services: services);
            shell.OutputWorkbench.SelectedSection = OutputWorkbenchSection.AssemblePdf;
            var model = shell.OutputWorkbench.Assembly;
            await model.AddFolderCommand.ExecuteAsync(null);
            Assert.Single(model.Sources);
            Assert.Equal(root.Name, model.Sources[0].Name);
            Assert.Equal("Folder", model.Sources[0].Kind);
            Assert.Equal(string.Empty, model.OutputPath);
            await model.ChooseOutputCommand.ExecuteAsync(null);
            await model.RunCommand.ExecuteAsync(null);
            Assert.True(model.HasOutput, model.Summary);
            Assert.Equal(3, PdfDocument.Load(model.PublishedPath!).Inspect().PageCount);
            Assert.Equal(first.Reads, first.ClosedReads);
            Assert.Equal(second.Reads, second.ClosedReads);
            Assert.Equal(1, first.Disposals);
            Assert.Equal(1, second.Disposals);
            Assert.Equal(1, nested.Disposals);
            Assert.Equal(0, root.Disposals);
            Assert.True(root.Enumerations >= 3);
            var window = new Window { Width = width, Height = height, Content = new OutputIntakeWorkbenchView { DataContext = shell } };
            try {
                window.Show(); window.UpdateLayout();
                using var frame = window.CaptureRenderedFrame();
                Assert.NotNull(frame);
                string? visual = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrWhiteSpace(visual)) {
                    Directory.CreateDirectory(visual);
                    frame.Save(Path.Combine(visual, $"provider-folder-{width}-{(dark ? "dark" : "light")}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task CancelledFolderSelectionReleasesEverySelectedItem() {
        using var storage = new StudioStorageAccess();
        var first = new Folder("content://folder/first", "First", []);
        var second = new Folder("content://folder/second", "Second", []);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => storage.RegisterFolderAsync([first.Item, second.Item], cancellation.Token));
        Assert.Equal(1, first.Disposals); Assert.Equal(1, second.Disposals);
    }

    private sealed class Folder {
        internal IStorageFolder Item { get; }
        internal string Name { get; }
        internal int Disposals;
        internal int Enumerations;
        internal Folder(string location, string name, IReadOnlyList<IStorageItem> children) {
            Name = name;
            Item = DispatchProxy.Create<IStorageFolder, TestStorageFile.StorageProxy>();
            ((TestStorageFile.StorageProxy)(object)Item).Call = (method, _) => method switch {
                "get_Name" => name, "get_Path" => new Uri(location), "get_CanBookmark" => false,
                "GetItemsAsync" => Enumerate(), "Dispose" => Dispose(), _ => throw new NotSupportedException(method)
            };
            async IAsyncEnumerable<IStorageItem> Enumerate() {
                Enumerations++; await Task.CompletedTask;
                foreach (var child in children) yield return child;
            }
            object? Dispose() { Disposals++; return null; }
        }
    }
}

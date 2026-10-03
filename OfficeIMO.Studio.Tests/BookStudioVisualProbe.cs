using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

/// <summary>Opt-in native publishing acceptance using disposable files and a captured rendered window.</summary>
internal static class BookStudioVisualProbe {
    internal static int Run(string outputRoot, string state, int width, int height) {
        if (state is not ("details" or "chapters" or "preview" or "review" or "empty") || width < 700 || height < 600) return 2;
        string root = Path.GetFullPath(outputRoot);
        Directory.CreateDirectory(root);
        var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "profile")));
        return AppBuilder.Configure(() => new App(services)).UsePlatformDetect().LogToTrace()
            .AfterSetup(_ => Dispatcher.UIThread.Post(async () => {
                var lifetime = (IClassicDesktopStyleApplicationLifetime)Application.Current!.ApplicationLifetime!;
                try {
                    var window = (MainWindow)lifetime.MainWindow!;
                    window.Width = width; window.Height = height;
                    window.Title = "EPUB publishing qualification";
                    window.ViewModel.ShowBookWorkbenchCommand.Execute(null);
                    var book = window.ViewModel.BookWorkbench;
                    if (state != "empty") {
                        string source = Path.Combine(root, "manuscript.md");
                        await File.WriteAllTextAsync(source, "---\ntitle: A field guide to publishing\nauthor: Sample author\nlanguage: en\n---\n# Opening chapter\n\nA reflowable book keeps its **meaning** as the reader changes the page size.\n\n- [x] Import the manuscript\n- [ ] Review the final EPUB\n\n## A small table\n\n| Stage | Result |\n| --- | --- |\n| Import | Typed book |\n| Export | EPUB file |\n\n# Second chapter\n\nA second chapter with a [return link](#opening-chapter).\n");
                        await book.OpenLocationAsync(source, default);
                        book.BookTitle = "A field guide to EPUB publishing";
                        await book.ApplyEditsCommand.ExecuteAsync(null);
                        string project = Path.Combine(root, "book.oibook"), epub = Path.Combine(root, "book.epub");
                        window.ViewModel.FileDialogs = new LocalDialogs(window, services.Storage, project, epub);
                        await book.SaveProjectCommand.ExecuteAsync(null);
                        if (book.IsDirty) throw new InvalidOperationException(book.Status);
                        await book.ExportBookCommand.ExecuteAsync(null);
                        BookProject restored = BookProject.LoadProject(await File.ReadAllBytesAsync(project));
                        using var exported = File.OpenRead(epub);
                        var publication = OfficeIMO.Epub.EpubDocument.Load(exported);
                        if (restored.Publication.Title != book.BookTitle || publication.Title != book.BookTitle || publication.Chapters.Count != 2)
                            throw new InvalidDataException("Native publication acceptance failed.");
                        if (state == "preview") await book.PreviewChapterCommand.ExecuteAsync(null);
                    }
                    var tabs = window.GetVisualDescendants().OfType<TabControl>().First(control => control.DataContext is BookWorkbenchViewModel);
                    tabs.SelectedIndex = state switch { "chapters" => 1, "preview" => 2, "review" => 3, _ => 0 };
                    await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Render);
                    // Let the native compositor commit the newly selected content before rendering its visual tree.
                    await Task.Delay(300);
                    var pixels = new PixelSize((int)Math.Ceiling(window.Bounds.Width * window.RenderScaling), (int)Math.Ceiling(window.Bounds.Height * window.RenderScaling));
                    using var capture = new RenderTargetBitmap(pixels, new Vector(96 * window.RenderScaling, 96 * window.RenderScaling));
                    capture.Render(window); capture.Save(Path.Combine(root, state + "-" + width + "x" + height + ".png"), PngBitmapEncoderOptions.Default);
                    await File.WriteAllTextAsync(Path.Combine(root, state + "-status.txt"), book.Status + "\n" + book.PreviewStatus + "\n" + string.Join("\n", book.Diagnostics));
                    Console.WriteLine("VERIFIED " + state + " " + width + "x" + height);
                    lifetime.Shutdown(0);
                } catch (Exception error) { Console.Error.WriteLine(error); lifetime.Shutdown(1); }
            }, DispatcherPriority.Background)).StartWithClassicDesktopLifetime([]);
    }
    private sealed class LocalDialogs(MainWindow window, StudioStorageAccess storage, string project, string epub) : IStudioFileDialogs {
        public Task<string?> PickOpenFileAsync(string title, StudioFileType type, CancellationToken token) => Task.FromResult<string?>(null);
        public async Task<string?> PickSaveFileAsync(string title, string suggestedName, StudioFileType type, CancellationToken token) {
            string path = type.Extensions.Contains("oibook") ? project : epub;
            if (!File.Exists(path)) await File.WriteAllBytesAsync(path, [], token);
            var item = await window.StorageProvider.TryGetFileFromPathAsync(new Uri(path)) ?? throw new IOException("Native provider file unavailable.");
            return await storage.RegisterAsync(item, token);
        }
        public Task<string?> PickFolderAsync(string title, CancellationToken token) => Task.FromResult<string?>(null);
    }
}

using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

/// <summary>Opt-in native Studio window using synthetic documents and an isolated profile for UI acceptance.</summary>
internal static class StudioVisualProbe {
    internal static int Run(string outputRoot) {
        string root = Path.GetFullPath(outputRoot);
        Directory.CreateDirectory(root);
        string documentPath = Path.Combine(root, "redaction-review.pdf");
        if (!File.Exists(documentPath)) {
            PdfDocument.Create(compose => {
                compose.Page(page => page.Size(600, 800).Content(content => {
                    content.Item(item => item.Paragraph(paragraph => paragraph.Text("Quarterly account review")));
                    content.Item(item => item.Paragraph(paragraph => paragraph.Text("Private account 123-456")));
                    content.Item(item => item.Paragraph(paragraph => paragraph.Text("This public summary should remain readable.")));
                }));
                compose.Page(page => page.Size(600, 800).Content(content => {
                    content.Item(item => item.Paragraph(paragraph => paragraph.Text("Second account")));
                    content.Item(item => item.Paragraph(paragraph => paragraph.Text("Private account 789-012")));
                }));
            }).Save(documentPath);
        }
        var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "profile")));
        return AppBuilder.Configure(() => new App(services)).UsePlatformDetect().LogToTrace()
            .AfterSetup(_ => Dispatcher.UIThread.Post(async () => {
                try {
                    var desktop = (IClassicDesktopStyleApplicationLifetime)Application.Current!.ApplicationLifetime!;
                    var window = (MainWindow)desktop.MainWindow!;
                    window.Width = 1280;
                    window.Height = 900;
                    await window.TabHost.OpenDocumentAsync(documentPath);
                    window.ViewModel.ShowProtectModeCommand.Execute(null);
                    window.ViewModel.RedactionSearchText = "Private account";
                    Console.WriteLine("READY " + documentPath);
                    Console.Out.Flush();
                } catch (Exception error) {
                    Console.Error.WriteLine(error);
                }
            }, DispatcherPriority.Background))
            .StartWithClassicDesktopLifetime([]);
    }
}

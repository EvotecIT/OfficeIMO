using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Assistant;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

/// <summary>Repeatable native UI acceptance against synthetic documents and an isolated profile.</summary>
internal static class StudioExperienceProbe {
    internal static int Run(string root, string scenario, int width, int height, string culture, string theme) {
        if (scenario is not ("home" or "document" or "tabs" or "assistant" or "connections" or "ocr" or "convert")
            || width < 960 || height < 620 || !Enum.TryParse(theme, true, out StudioThemePreference appearance)) return 2;
        root = Path.GetFullPath(root);
        Directory.CreateDirectory(root);
        var paths = new StudioDataPaths(Path.Combine(root, "profile"));
        new JsonStudioPreferencesStore(paths.PreferencesPath).Save(new StudioPreferences { UiCulture = culture, Theme = appearance, RememberSession = false });
        var services = StudioApplicationServices.Create(paths);
        string source = Path.Combine(root, "quarterly-review.pdf");
        PdfDocument.Create(document => document.Page(page => page.Content(content => {
            content.Text("Quarterly review");
            content.Text("The approved budget is 42,000 EUR. The delivery deadline is 30 September.");
            content.Text("Action: prepare the project summary. Owner: the delivery team.");
        }))).Save(source);
        return AppBuilder.Configure(() => new App(services)).UsePlatformDetect().LogToTrace()
            .AfterSetup(builder => Dispatcher.UIThread.Post(async () => {
                try {
                    var lifetime = (IClassicDesktopStyleApplicationLifetime)Application.Current!.ApplicationLifetime!;
                    var window = (MainWindow)lifetime.MainWindow!;
                    window.Width = width; window.Height = height;
                    if (scenario != "home") await window.TabHost.OpenDocumentAsync(source);
                    if (scenario == "tabs") {
                        string second = Path.Combine(root, "delivery-notes.pdf");
                        PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("Delivery notes: the review is complete.")))).Save(second);
                        await window.TabHost.OpenDocumentAsync(second);
                    }
                    if (scenario is "assistant" or "connections") window.ViewModel.ToggleAssistantCommand.Execute(null);
                    if (scenario == "ocr") window.ViewModel.ShowOcrCommand.Execute(null);
                    if (scenario == "convert") window.ViewModel.Commands["Convert"].Execute(null);
                    if (scenario == "connections") {
                        var connections = new ConnectionsWindow { DataContext = services.AiConnections };
                        _ = connections.ShowDialog<bool>(window);
                    }
                    Console.WriteLine($"READY {scenario} {width}x{height} {culture} {theme}");
                    Console.Out.Flush();
                } catch (Exception error) { Console.Error.WriteLine(error); }
            }, DispatcherPriority.Background)).StartWithClassicDesktopLifetime([]);
    }
}

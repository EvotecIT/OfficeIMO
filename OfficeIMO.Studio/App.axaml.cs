using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio;

public sealed partial class App : Application {
    public override void Initialize() {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted() {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop) {
            var window = new MainWindow();
            window.OpenInitialDocument(desktop.Args);
            desktop.MainWindow = window;
        }

        base.OnFrameworkInitializationCompleted();
    }
}

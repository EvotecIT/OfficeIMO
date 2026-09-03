using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Threading;
using Avalonia.Markup.Xaml;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Diagnostics;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio;

public sealed partial class App : Application {
    private bool _diagnosticHandlersAttached;

    internal StudioApplicationServices Services { get; private set; } = null!;

    public override void Initialize() {
        Services ??= StudioApplicationServices.CreateDefault();
        StudioLocalization.Configure(Services.Localizer);
        AvaloniaXamlLoader.Load(this);
        RequestedThemeVariant = Services.Preferences.Current.Theme switch {
            StudioThemePreference.Light => Avalonia.Styling.ThemeVariant.Light,
            StudioThemePreference.Dark => Avalonia.Styling.ThemeVariant.Dark,
            StudioThemePreference.HighContrast => StudioThemeVariants.HighContrast,
            _ => Avalonia.Styling.ThemeVariant.Default
        };
        Services.Preferences.Changed += OnPreferencesChanged;
    }

    private void OnPreferencesChanged(object? sender, EventArgs eventArgs) {
        RequestedThemeVariant = Services.Preferences.Current.Theme switch {
            StudioThemePreference.Light => Avalonia.Styling.ThemeVariant.Light,
            StudioThemePreference.Dark => Avalonia.Styling.ThemeVariant.Dark,
            StudioThemePreference.HighContrast => StudioThemeVariants.HighContrast,
            _ => Avalonia.Styling.ThemeVariant.Default
        };
    }

    public override void OnFrameworkInitializationCompleted() {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop) {
            AttachDiagnosticHandlers();
            var window = new MainWindow(Services);
            window.OpenInitialDocument(desktop.Args);
            desktop.MainWindow = window;
        }

        base.OnFrameworkInitializationCompleted();
    }

    private void AttachDiagnosticHandlers() {
        if (_diagnosticHandlersAttached) return;
        _diagnosticHandlersAttached = true;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
        Dispatcher.UIThread.UnhandledException += OnDispatcherUnhandledException;
        Services.Diagnostics.Write(StudioDiagnosticLevel.Information, "Application", "Started");
    }

    private void OnUnhandledException(object sender, UnhandledExceptionEventArgs eventArgs) {
        Services.Diagnostics.Write(
            StudioDiagnosticLevel.Critical,
            "Application",
            "UnhandledException",
            eventArgs.ExceptionObject as Exception);
    }

    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs eventArgs) {
        Services.Diagnostics.Write(StudioDiagnosticLevel.Error, "Application", "UnobservedTaskException", eventArgs.Exception);
    }

    private void OnDispatcherUnhandledException(object? sender, DispatcherUnhandledExceptionEventArgs eventArgs) {
        Services.Diagnostics.Write(StudioDiagnosticLevel.Critical, "UserInterface", "UnhandledException", eventArgs.Exception);
    }
}

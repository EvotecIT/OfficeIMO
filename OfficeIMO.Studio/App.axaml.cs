using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Threading;
using Avalonia.Markup.Xaml;
using Avalonia.Styling;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Diagnostics;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio;

public sealed partial class App : Application {
    private bool _diagnosticHandlersAttached;

    public App() { }

    internal App(StudioApplicationServices services) => Services = services;

    internal StudioApplicationServices Services { get; private set; } = null!;

    internal Func<App, Avalonia.Controls.Control>? SingleViewFactory { get; init; }

    /// <summary>The desktop instance lock whose forwarded launches open in the main window.</summary>
    internal StudioSingleInstance? SingleInstance { get; init; }

    public override void Initialize() {
        Services ??= StudioApplicationServices.CreateDefault();
        StudioLocalization.Configure(Services.Localizer);
        AvaloniaXamlLoader.Load(this);
        InitializeNativeApplicationMenu();
        ApplyLocationPolicy(Styles, StudioDistributionPolicy.RequiresPickedLocations);
        ApplyFlowDirection(Styles, Services.Localizer.Culture);
        ApplyDensity();
        RequestedThemeVariant = Services.Preferences.Current.Theme switch {
            StudioThemePreference.Light => Avalonia.Styling.ThemeVariant.Light,
            StudioThemePreference.Dark => Avalonia.Styling.ThemeVariant.Dark,
            StudioThemePreference.HighContrast => StudioThemeVariants.HighContrast,
            _ => Avalonia.Styling.ThemeVariant.Default
        };
        Services.Preferences.Changed += OnPreferencesChanged;
    }

    private void OnPreferencesChanged(object? sender, EventArgs eventArgs) {
        ApplyDensity();
        RequestedThemeVariant = Services.Preferences.Current.Theme switch {
            StudioThemePreference.Light => Avalonia.Styling.ThemeVariant.Light,
            StudioThemePreference.Dark => Avalonia.Styling.ThemeVariant.Dark,
            StudioThemePreference.HighContrast => StudioThemeVariants.HighContrast,
            _ => Avalonia.Styling.ThemeVariant.Default
        };
    }

    /// <summary>
    /// Where user files are reachable only through a picker, location boxes (class <c>path</c>) show the chosen location
    /// but cannot be typed into: a typed path would bypass the permission the picker grants.
    /// </summary>
    internal static void ApplyLocationPolicy(Styles styles, bool requiresPickedLocations) {
        if (!requiresPickedLocations) return;
        styles.Add(new Style(selector => selector!.OfType<Avalonia.Controls.TextBox>().Class("path")) {
            Setters = { new Setter(Avalonia.Controls.TextBox.IsReadOnlyProperty, true) }
        });
    }

    /// <summary>
    /// A right-to-left interface language mirrors every window, menu and dialog. Documents keep their own direction:
    /// PDF pages, scans, signatures, previews and images are never mirrored, and file locations read left to right.
    /// </summary>
    internal static void ApplyFlowDirection(Styles styles, System.Globalization.CultureInfo culture) {
        if (!culture.TextInfo.IsRightToLeft) return;
        styles.Add(new Style(selector => selector!.Is<Avalonia.Controls.TopLevel>()) {
            Setters = { new Setter(Visual.FlowDirectionProperty, Avalonia.Media.FlowDirection.RightToLeft) }
        });
        styles.Add(new Style(selector => Selectors.Or(
            selector!.OfType<Features.Reader.PdfPageCanvas>(),
            selector!.OfType<Features.Workflows.ScanSelectionCanvas>(),
            selector!.OfType<Features.Sign.SignaturePad>(),
            selector!.OfType<Features.Editor.TextEditPreview>(),
            selector!.OfType<Avalonia.Controls.Image>(),
            selector!.OfType<Avalonia.Controls.TextBox>().Class("path"))) {
            Setters = { new Setter(Visual.FlowDirectionProperty, Avalonia.Media.FlowDirection.LeftToRight) }
        });
    }

    private void ApplyDensity() {
        bool compact = Services.Preferences.Current.Density == StudioDensityPreference.Compact;
        Resources["StudioBodyFontSize"] = compact ? 13D : 14D;
        Resources["StudioControlHeight"] = compact ? 32D : 38D;
        Resources["StudioControlPadding"] = compact ? new Thickness(8, 3) : new Thickness(10, 6);
        Resources["StudioModePadding"] = compact ? new Thickness(12, 5) : new Thickness(14, 8);
    }

    public override void OnFrameworkInitializationCompleted() {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop) {
            AttachDiagnosticHandlers();
            var window = new MainWindow(Services);
            window.OpenInitialDocument(desktop.Args);
            desktop.MainWindow = window;
            AttachFileActivation(desktop, window);
            AttachForwardedLaunches(window);
        } else if (ApplicationLifetime is ISingleViewApplicationLifetime singleView && SingleViewFactory is not null) {
            AttachDiagnosticHandlers();
            singleView.MainView = SingleViewFactory(this);
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
        if (!CanRecoverFromUserInterfaceFailure(eventArgs.Exception, DateTime.UtcNow)) return;
        // Keep open documents and unsaved work reachable; the notice asks the user to save and restart.
        eventArgs.Handled = true;
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime { MainWindow: MainWindow window }) window.ReportUnexpectedError();
    }

    private readonly Queue<DateTime> _recentUserInterfaceFailures = new();

    // A failure that repeats on every frame or input would otherwise loop forever; let a burst terminate.
    internal bool CanRecoverFromUserInterfaceFailure(Exception error, DateTime now) {
        if (error is OutOfMemoryException or InsufficientExecutionStackException or AccessViolationException) return false;
        while (_recentUserInterfaceFailures.Count > 0 && now - _recentUserInterfaceFailures.Peek() > TimeSpan.FromSeconds(10))
            _recentUserInterfaceFailures.Dequeue();
        _recentUserInterfaceFailures.Enqueue(now);
        return _recentUserInterfaceFailures.Count <= 5;
    }
}

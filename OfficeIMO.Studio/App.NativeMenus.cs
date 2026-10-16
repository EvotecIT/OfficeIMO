using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio;

public sealed partial class App {
    /// <summary>
    /// The macOS application menu. Setting it replaces Avalonia's default About entry; AppKit adds Services, Hide and Quit.
    /// The entries act on the main window, which owns About and Settings.
    /// </summary>
    private void InitializeNativeApplicationMenu() {
        if (!OperatingSystem.IsMacOS()) return;
        void OnMainWindow(Action<MainWindow> action) {
            if (ApplicationLifetime is not IClassicDesktopStyleApplicationLifetime { MainWindow: MainWindow window }) return;
            window.Activate();
            action(window);
        }
        var menu = new NativeMenu();
        foreach (StudioMenuEntry entry in new StudioMenuEntry[] {
                     new StudioMenuCommand("About", Services.Localizer.Get("Apple.About"), new RelayCommand(() => OnMainWindow(window => window.ShowAbout()))),
                     StudioMenuSeparator.Instance,
                     new StudioMenuCommand("Settings", Services.Localizer.Get("Apple.Settings"),
                         new RelayCommand(() => OnMainWindow(window => window.ViewModel.Commands["Settings"].Execute(null))), StudioShortcuts.Primary(Avalonia.Input.Key.OemComma))
                 })
            menu.Items.Add(StudioMenuRenderer.ToNative(entry));
        NativeMenu.SetMenu(this, menu);
    }
}

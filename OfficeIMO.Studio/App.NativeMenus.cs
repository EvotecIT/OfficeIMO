using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio;

public sealed partial class App {
    private void InitializeNativeApplicationMenu() {
        if (!OperatingSystem.IsMacOS()) return;
        var menu = new NativeMenu();
        menu.Items.Add(new NativeMenuItem(Services.Localizer.Get("Apple.Settings")) {
            Gesture = KeyGesture.Parse("Meta+OemComma"),
            Command = new RelayCommand(() => {
                if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime { MainWindow: MainWindow window }) {
                    window.Activate();
                    window.ViewModel.Commands["Settings"].Execute(null);
                }
            })
        });
        NativeMenu.SetMenu(this, menu);
    }
}

using Avalonia;
using Avalonia.Controls;
using Avalonia.Platform;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindow {
    private StudioWindowPlacement? _normalPlacement;
    private bool _restoreMaximized;
    private WindowState _stateBeforeMinimize = WindowState.Normal;

    // Reopen where the user left the window; fall back to a size that fits the current screen.
    private void RestoreWindowPlacement() {
        PositionChanged += (_, _) => RememberNormalPlacement();
        PropertyChanged += (_, change) => {
            if (change.Property == WindowStateProperty && WindowState != WindowState.Minimized) _stateBeforeMinimize = WindowState;
            if (change.Property == ClientSizeProperty || change.Property == WindowStateProperty) RememberNormalPlacement();
        };
        Opened += (_, _) => {
            if (_restoreMaximized) WindowState = WindowState.Maximized;
        };

        IReadOnlyList<Screen> screens = Screens.All;
        if (screens.Count == 0) return;
        StudioWindowPlacement? saved = _services.Preferences.Current.MainWindow;
        Screen? target = saved is null ? null : screens.FirstOrDefault(screen => Visible(screen, saved));
        if (saved is not null && target is not null) {
            Width = Math.Min(saved.Width, target.WorkingArea.Width / target.Scaling);
            Height = Math.Min(saved.Height, target.WorkingArea.Height / target.Scaling);
            WindowStartupLocation = WindowStartupLocation.Manual;
            int width = (int)Math.Ceiling(Width * target.Scaling);
            int height = (int)Math.Ceiling(Height * target.Scaling);
            Position = new PixelPoint(
                Math.Clamp(saved.X, target.WorkingArea.X, Math.Max(target.WorkingArea.X, target.WorkingArea.Right - width)),
                Math.Clamp(saved.Y, target.WorkingArea.Y, Math.Max(target.WorkingArea.Y, target.WorkingArea.Bottom - height)));
            _restoreMaximized = saved.Maximized;
            return;
        }
        if ((Screens.Primary ?? screens[0]) is { } primary) {
            // The designed default is larger than small laptop screens; keep a margin around it there.
            Width = Math.Min(Width, primary.WorkingArea.Width / primary.Scaling * 0.92);
            Height = Math.Min(Height, primary.WorkingArea.Height / primary.Scaling * 0.92);
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }
    }

    // At least a title-bar-sized strip must land on a connected screen so the window can be moved.
    private static bool Visible(Screen screen, StudioWindowPlacement placement) {
        var titleBar = new PixelRect(placement.X, placement.Y, (int)Math.Max(1, placement.Width * screen.Scaling), (int)(44 * screen.Scaling));
        PixelRect overlap = screen.WorkingArea.Intersect(titleBar);
        return overlap.Width >= 120 * screen.Scaling && overlap.Height >= 20 * screen.Scaling;
    }

    private void RememberNormalPlacement() {
        if (WindowState != WindowState.Normal || !IsVisible) return;
        _normalPlacement = new StudioWindowPlacement { X = Position.X, Y = Position.Y, Width = ClientSize.Width, Height = ClientSize.Height };
    }

    private void SaveWindowPlacement() {
        StudioWindowPlacement? normal = _normalPlacement;
        if (normal is null) return;
        // Full screen is not restored at launch; the window returns to the state full screen was entered from.
        WindowState saved = WindowState == WindowState.Minimized ? _stateBeforeMinimize : WindowState;
        if (saved == WindowState.FullScreen) saved = _stateBeforeFullScreen;
        StudioWindowPlacement placement = normal with { Maximized = saved == WindowState.Maximized };
        _services.Preferences.UpdatePresentation(current => current with { MainWindow = placement });
    }
}

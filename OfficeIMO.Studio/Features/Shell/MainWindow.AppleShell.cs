using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Markup.Xaml.Styling;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindow {
    private bool _sidebarHidden;

    private void InitializeAppleShell() {
        Classes.Set("apple", OperatingSystem.IsMacOS());
        if (OperatingSystem.IsMacOS()) {
            Resources.MergedDictionaries.Add(new ResourceInclude(new Uri("avares://OfficeIMO.Studio/")) {
                Source = new Uri("avares://OfficeIMO.Studio/Features/Shell/AppleShellResources.axaml")
            });
            // Let AppKit own the traffic lights and title. The document toolbar stays below it.
            ExtendClientAreaToDecorationsHint = false;
            FontFamily = FontFamily.Default;
            TitleBarLogo.IsVisible = false;
        }
        RefreshNativeMenus();
        TabHost.PropertyChanged += (_, change) => {
            if (change.PropertyName is nameof(StudioDocumentTabHost.HasTabs) or nameof(StudioDocumentTabHost.SelectedTab))
                ApplyResponsiveLayout(Bounds.Width);
        };
        ApplyResponsiveLayout(Width);
    }

    private void OnSidebarToggleClick(object? sender, RoutedEventArgs e) {
        _sidebarHidden = !_sidebarHidden;
        ApplyResponsiveLayout(Bounds.Width);
    }

    private async void OnCloseDocumentClick(object? sender, RoutedEventArgs e) => await CloseSelectedDocumentTabAsync();

    /// <summary>Adapts shared navigation to the available content width, including narrow split views.</summary>
    internal void ApplyResponsiveLayout(double width) {
        if (width <= 0) return;
        bool narrow = width < 700;
        bool apple = OperatingSystem.IsMacOS();
        double sidebar = narrow ? 0 : apple ? (_sidebarHidden ? 0 : width >= 1100 ? 208 : 176) : 72;
        Classes.Set("narrow", narrow);
        ShellRoot.ColumnDefinitions[0].Width = new GridLength(sidebar);
        ShellRoot.RowDefinitions[0].Height = new GridLength(narrow ? 52 : 44);
        NavigationRail.IsVisible = sidebar > 0;
        CompactNavigation.IsVisible = narrow;
        SidebarToggle.IsVisible = apple && !narrow;
        Grid.SetColumnSpan(SidebarToggle, _sidebarHidden ? 2 : 1);
        TitleBarLogo.IsVisible = !apple && !narrow;
        Grid.SetColumn(TitleBar, sidebar > 0 ? 1 : 0);
        Grid.SetColumnSpan(TitleBar, sidebar > 0 ? 1 : 2);
        TitleBar.Margin = new Thickness(apple && _sidebarHidden && !narrow ? 52 : 0, 0, apple ? 10 : 0, 0);
        TitleBar.ColumnDefinitions[0].Width = new GridLength(1, GridUnitType.Star);
        TitleBar.ColumnDefinitions[1].Width = narrow && TabHost.HasTabs ? GridLength.Auto : new GridLength(0);
        CompactCloseDocumentButton.IsVisible = narrow && TabHost.HasTabs;
        TitleDocuments.IsVisible = !narrow;
        CompactDocumentPicker.IsVisible = narrow && TabHost.HasTabs;
        CompactTitle.IsVisible = narrow && !TabHost.HasTabs;
        CommandSearchLabel.IsVisible = !apple && !narrow && width >= 1100;
        CommandSearchKeycap.IsVisible = !apple && !narrow && width >= 1100;
        CommandSearchButton.Width = narrow ? 44 : apple || width < 1100 ? 36 : 240;
        CommandSearchButton.Padding = apple || narrow || width < 1100 ? new Thickness(10) : new Thickness(10, 0, 6, 0);
        ThemeToggle.IsVisible = !apple && !narrow;
        DocumentListButton.IsVisible = !narrow && TabHost.HasTabs;
        AssistantToggle.IsVisible = !narrow;
        // Open remains reachable when the tab strip gives way to the document picker.
        if (narrow) {
            TitleDocuments.IsVisible = true;
            DocumentTabs.IsVisible = false;
            AppTitleText.IsVisible = false;
            OpenDocumentTabButton.IsVisible = true;
            TitleDocuments.HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right;
            Grid.SetColumn(TitleDocuments, 3);
            OpenDocumentTabButton.Width = 44;
            OpenDocumentTabButton.Height = 44;
        } else {
            Grid.SetColumn(TitleDocuments, 0);
            TitleDocuments.HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Stretch;
            DocumentTabs.IsVisible = TabHost.HasTabs;
            AppTitleText.IsVisible = !TabHost.HasTabs;
            OpenDocumentTabButton.Width = 30;
            OpenDocumentTabButton.Height = 30;
        }
        DocumentTabs.VerticalAlignment = apple ? Avalonia.Layout.VerticalAlignment.Center : Avalonia.Layout.VerticalAlignment.Bottom;
        DocumentTabs.MaxWidth = double.PositiveInfinity;
        ContentSurface.CornerRadius = apple || narrow ? new CornerRadius(0) : new CornerRadius(10, 0, 0, 0);
        ContentSurface.BorderThickness = new Thickness(sidebar > 0 ? 1 : 0, 1, 0, 0);
        AssistantHost.OpenPaneLength = Math.Min(400, Math.Max(0, width - sidebar));
        AssistantHost.DisplayMode = width >= 1500 ? SplitViewDisplayMode.Inline : SplitViewDisplayMode.Overlay;
        IsCompactLayout = width < 1180;
        double workspaceWidth = Math.Max(0, width - sidebar - (width >= 1500 && AssistantHost.IsPaneOpen ? 400 : 0));
        DocumentWorkspace.ApplyResponsiveLayout(workspaceWidth);
        ConversionView.ApplyResponsiveLayout(workspaceWidth);
        DocumentHealthView.ApplyResponsiveLayout(workspaceWidth);
        RevealSelectedDocumentTab();
    }
}

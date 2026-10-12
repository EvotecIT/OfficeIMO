using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Styling;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Home;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>
/// The Windows and Linux application menu behind the title-bar logo. It is built from the shared command
/// catalog each time it opens, so labels, shortcuts, availability and unavailable reasons match the palette.
/// macOS uses the native menu bar instead.
/// </summary>
public sealed partial class MainWindow {
    internal const string DocumentationUri = "https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Studio";
    internal const string ReportProblemUri = "https://github.com/EvotecIT/OfficeIMO/issues";
    private MenuFlyout? _appMenu;

    private void InitializeAppMenu() {
        if (OperatingSystem.IsMacOS()) return;
        _appMenu = new MenuFlyout { Placement = PlacementMode.BottomEdgeAlignedLeft };
        // A menu flyout sizes its popup from the items present when it opens, so build them before showing.
        AppMenuButton.Click += (_, _) => OpenAppMenu();
        AppMenuButton.IsVisible = true;
        TitleBarLogo.IsVisible = false;
    }

    internal MenuFlyout? AppMenu => _appMenu;

    /// <summary>Opens the application menu, optionally with one top-level menu expanded (Alt+letter).</summary>
    internal bool OpenAppMenu(string? menuId = null) {
        if (_appMenu is null) return false;
        if (_appMenu.IsOpen) { _appMenu.Hide(); return true; }
        RebuildAppMenu();
        // Narrow layouts hide the logo; the menu still opens from the keyboard at the window corner.
        _appMenu.ShowAt(AppMenuButton.IsEffectivelyVisible ? AppMenuButton : ShellRoot);
        MenuItem? target = _appMenu.Items.OfType<MenuItem>().FirstOrDefault(item => menuId is null || Equals(item.Tag, menuId));
        if (target is null) return true;
        // Menu items attach to the popup after it opens; focus and expand them once they are in the visual tree.
        Avalonia.Threading.Dispatcher.UIThread.Post(() => {
            if (!_appMenu.IsOpen) return;
            target.Focus(NavigationMethod.Directional);
            if (menuId is null) return;
            target.Open();
            Avalonia.Threading.Dispatcher.UIThread.Post(() => {
                if (target.IsSubMenuOpen) target.Items.OfType<MenuItem>().FirstOrDefault(item => item.IsEnabled)?.Focus(NavigationMethod.Directional);
            }, Avalonia.Threading.DispatcherPriority.Loaded);
        }, Avalonia.Threading.DispatcherPriority.Loaded);
        return true;
    }

    private void RebuildAppMenu() {
        if (_appMenu is null) return;
        _appMenu.Items.Clear();
        foreach (MenuItem menu in BuildAppMenu()) _appMenu.Items.Add(menu);
    }

    internal IReadOnlyList<MenuItem> BuildAppMenu() {
        StudioCommandCatalog commands = ViewModel.Commands;
        string Text(string key) => _services.Localizer.Get("Menu." + key);

        MenuItem Top(string id, params object[] items) {
            var menu = new MenuItem { Header = AccessKeys()[id].Header, Tag = id };
            foreach (object item in items) menu.Items.Add(item);
            return menu;
        }

        MenuItem Command(string id) {
            StudioCommandItem command = commands[id];
            var item = new MenuItem { Header = command.Title, Command = command, Tag = id, InputGesture = command.Gesture };
            if (command.UnavailableReason is { Length: > 0 } reason) {
                ToolTip.SetTip(item, reason);
                ToolTip.SetShowOnDisabled(item, true);
            }
            return item;
        }

        MenuItem Action(string key, Action action, KeyGesture? gesture = null, bool enabled = true) => new() {
            Header = Text(key), Tag = key, Command = new RelayCommand(action, () => enabled), InputGesture = gesture
        };

        MenuItem AsyncAction(string key, Func<Task> action, KeyGesture? gesture = null, bool enabled = true) => new() {
            Header = Text(key), Tag = key, Command = new AsyncRelayCommand(action, () => enabled), InputGesture = gesture
        };

        MenuItem recent = new() { Header = Text("OpenRecent"), Tag = "OpenRecent" };
        foreach (RecentDocumentViewModel recentDocument in ViewModel.RecentDocuments.Take(10)) {
            var entry = new MenuItem { Header = recentDocument.FileName, Command = ViewModel.OpenRecentCommand, CommandParameter = recentDocument };
            ToolTip.SetTip(entry, recentDocument.Path);
            recent.Items.Add(entry);
        }
        if (recent.Items.Count > 0) {
            recent.Items.Add(new Separator());
            recent.Items.Add(new MenuItem { Header = Text("ClearRecent"), Command = ViewModel.ClearRecentDocumentsCommand });
        }
        recent.IsEnabled = recent.Items.Count > 0;

        MenuItem file = Top("File",
            Command("Open"), recent, new Separator(),
            Command("Save"), Command("SaveAs"), new Separator(),
            Command("Print"), new Separator(),
            AsyncAction("CloseTab", CloseSelectedDocumentTabAsync, StudioShortcuts.CloseTab, TabHost.HasTabs),
            AsyncAction("ReopenTab", TabHost.ReopenClosedTabAsync, StudioShortcuts.ReopenTab, TabHost.CanReopenClosedTab),
            new Separator(), Command("Settings"), new Separator(),
            Action("Exit", Close, StudioShortcuts.Quit));

        MenuItem edit = Top("Edit",
            Command("Undo"), Command("Redo"), new Separator(),
            Action("Find", FocusDocumentSearch, StudioShortcuts.Find, ViewModel.HasDocument),
            AsyncAction("Commands", ShowCommandPaletteAsync, StudioShortcuts.CommandSearch));

        MenuItem view = Top("View",
            Command("Home"), Command("Tools"), Command("Jobs"), new Separator(),
            Command("ZoomIn"), Command("ZoomOut"),
            Command("ActualSize"), Command("FitPage"), Command("FitWidth"), new Separator(),
            Command("FocusReading"), Command("PageNightMode"), Command("CompareSideBySide"), Command("Assistant"),
            Action("ToggleSidebar", () => OnSidebarToggleClick(null, new()), enabled: ViewModel.HasDocument),
            Action(WindowState == WindowState.FullScreen ? "LeaveFullScreen" : "FullScreen", ToggleFullScreen, StudioShortcuts.FullScreen),
            new Separator(), AppearanceMenu());

        MenuItem document = Top("Document",
            Command("Read"), Command("Comment"), Command("Edit"), Command("Pages"), Command("Forms"), Command("Protect"),
            new Separator(), Command("Watermark"), Command("PageNumbers"), Command("FlattenAnnotations"),
            new Separator(), Command("Split"),
            new Separator(), Command("Redact"), Command("ValidateSignatures"),
            new Separator(), Command("PreviousFormField"), Command("NextFormField"));

        // Tools follow the catalog's categories so the menu and the Tools page group them the same way.
        var tools = Top("ToolsMenu");
        foreach (StudioCommandGroup group in commands.GroupTools(null)) {
            var category = new MenuItem { Header = group.Title };
            foreach (StudioCommandItem tool in group.Items) category.Items.Add(Command(tool.Id));
            tools.Items.Add(category);
        }

        MenuItem window = Top("Window",
            Action("NextTab", () => TabHost.SelectRelativeTab(false), StudioShortcuts.NextTab, TabHost.Tabs.Count > 1),
            Action("PreviousTab", () => TabHost.SelectRelativeTab(true), StudioShortcuts.PreviousTab, TabHost.Tabs.Count > 1),
            Action("Minimize", () => WindowState = WindowState.Minimized));
        if (TabHost.HasTabs) {
            window.Items.Add(new Separator());
            foreach (StudioDocumentTabViewModel tab in TabHost.Tabs) {
                window.Items.Add(new MenuItem {
                    Header = tab.DisplayTitle, ToggleType = MenuItemToggleType.Radio,
                    IsChecked = ReferenceEquals(tab, TabHost.SelectedTab),
                    Command = new RelayCommand(() => { if (TabHost.Tabs.Contains(tab)) TabHost.SelectedTab = tab; })
                });
            }
        }

        MenuItem help = Top("Help",
            AsyncAction("KeyboardShortcuts", ShowCommandPaletteAsync),
            AsyncAction("Documentation", () => OpenUriAsync(new Uri(DocumentationUri))),
            AsyncAction("ReportProblem", () => OpenUriAsync(new Uri(ReportProblemUri))),
            new Separator(), Action("About", () => ViewModel.Commands["Settings"].Execute(null)));

        return [file, edit, view, document, tools, window, help];
    }

    private MenuItem AppearanceMenu() {
        string Text(string key) => _services.Localizer.Get("Settings.Theme" + key);
        StudioThemePreference current = _services.Preferences.Current.Theme;
        var menu = new MenuItem { Header = _services.Localizer.Get("Settings.Appearance"), Tag = "Appearance" };
        foreach ((StudioThemePreference preference, string key) in new[] {
                     (StudioThemePreference.System, "System"), (StudioThemePreference.Light, "Light"),
                     (StudioThemePreference.Dark, "Dark"), (StudioThemePreference.HighContrast, "HighContrast") }) {
            menu.Items.Add(new MenuItem {
                Header = Text(key), ToggleType = MenuItemToggleType.Radio, IsChecked = current == preference,
                Command = new RelayCommand(() => _services.Preferences.UpdatePresentation(value => value with { Theme = preference }))
            });
        }
        return menu;
    }

    private WindowState _stateBeforeFullScreen = WindowState.Normal;

    // Leaving full screen returns to the state it was entered from, so a maximized window stays maximized.
    private void ToggleFullScreen() {
        if (WindowState == WindowState.FullScreen) {
            WindowState = _stateBeforeFullScreen;
            return;
        }
        _stateBeforeFullScreen = WindowState == WindowState.Minimized ? WindowState.Normal : WindowState;
        WindowState = WindowState.FullScreen;
    }

    // F10 opens the menu and Alt+letter opens one menu, matching Windows and most Linux desktops.
    private bool TryOpenAppMenuFromKeyboard(KeyEventArgs e) {
        if (_appMenu is null) return false;
        if (StudioShortcuts.Matches(StudioShortcuts.ApplicationMenu, e.Key, e.KeyModifiers)) return OpenAppMenu();
        if (e.KeyModifiers != KeyModifiers.Alt || e.Key is < Key.A or > Key.Z) return false;
        char letter = (char)('A' + (e.Key - Key.A));
        string? menu = AccessKeys().FirstOrDefault(pair => pair.Value.Key == letter).Key;
        return menu is not null && OpenAppMenu(menu);
    }

    private static readonly string[] TopLevelMenuIds = ["File", "Edit", "View", "Document", "ToolsMenu", "Window", "Help"];

    /// <summary>
    /// One access key per top-level menu in the current language: the first letter of each name that an earlier menu has
    /// not taken, compared without accents so "Édition" answers to E. The underlined letter is the key that opens it.
    /// </summary>
    internal IReadOnlyDictionary<string, (char? Key, string Header)> AccessKeys() =>
        AssignAccessKeys(TopLevelMenuIds.Select(id => (id, _services.Localizer.Get("Menu." + id))));

    internal static IReadOnlyDictionary<string, (char? Key, string Header)> AssignAccessKeys(IEnumerable<(string Id, string Title)> menus) {
        var taken = new HashSet<char>();
        var result = new Dictionary<string, (char? Key, string Header)>(StringComparer.Ordinal);
        foreach ((string id, string title) in menus) {
            string plain = title.Replace("_", "__", StringComparison.Ordinal);
            (char? key, string header) = (null, plain);
            for (int index = 0; index < title.Length; index++) {
                char letter = BaseLetter(title[index]);
                if (letter is < 'A' or > 'Z' || !taken.Add(letter)) continue;
                key = letter;
                header = title[..index].Replace("_", "__", StringComparison.Ordinal) + "_" + title[index..].Replace("_", "__", StringComparison.Ordinal);
                break;
            }
            result[id] = (key, header);
        }
        return result;
    }

    private static char BaseLetter(char character) {
        string decomposed = character.ToString().Normalize(System.Text.NormalizationForm.FormD);
        return char.ToUpperInvariant(decomposed[0]);
    }
}

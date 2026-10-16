using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Styling;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Home;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>
/// The Windows and Linux application menu behind the title-bar logo. It is built from the shared menu definition each
/// time it opens, so labels, shortcuts, availability and unavailable reasons match the palette and the macOS menu bar.
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

    /// <summary>Renders the shared menu definition as the in-window menu, with an access key on each top-level menu.</summary>
    internal IReadOnlyList<MenuItem> BuildAppMenu() {
        IReadOnlyDictionary<string, (char? Key, string Header)> keys = AccessKeys();
        return BuildMenuModel(nativeMac: false).Select(group => (MenuItem)StudioMenuRenderer.ToMenuItem(group, keys[group.Id].Header)).ToArray();
    }

    /// <summary>
    /// The one menu definition. The macOS form follows the platform's conventions: About, Settings and Quit live in the
    /// application menu, Edit also drives text fields, and titles use the platform's wording.
    /// </summary>
    internal IReadOnlyList<StudioMenuGroup> BuildMenuModel(bool nativeMac) {
        StudioCommandCatalog commands = ViewModel.Commands;
        string Text(string key) => _services.Localizer.Get((nativeMac ? "Apple." : "Menu.") + key);
        StudioMenuSeparator separator = StudioMenuSeparator.Instance;

        StudioMenuGroup Group(string id, params StudioMenuEntry[] items) => new(id, Text(id), items);

        StudioMenuCommand Command(string id) {
            StudioCommandItem command = commands[id];
            // The macOS menu names commands that ask for more input with an ellipsis.
            string title = nativeMac && id is "Open" or "SaveAs" or "Settings" ? Text(id) : command.Title;
            return new StudioMenuCommand(id, title, command, command.Gesture) { ToolTip = command.UnavailableReason };
        }

        // Guards are evaluated when asked, so a menu kept open by the platform never acts on a stale state.
        StudioMenuCommand Action(string key, Action action, KeyGesture? gesture = null, Func<bool>? canExecute = null) =>
            new(key, Text(key), canExecute is null ? new RelayCommand(action) : new RelayCommand(action, canExecute), gesture);

        StudioMenuCommand AsyncAction(string key, Func<Task> action, KeyGesture? gesture = null, Func<bool>? canExecute = null) =>
            new(key, Text(key), canExecute is null ? new AsyncRelayCommand(action) : new AsyncRelayCommand(action, canExecute), gesture);

        var recentItems = new List<StudioMenuEntry>();
        foreach (RecentDocumentViewModel recentDocument in ViewModel.RecentDocuments.Take(10))
            recentItems.Add(new StudioMenuCommand("Recent", recentDocument.FileName, ViewModel.OpenRecentCommand) { Parameter = recentDocument, ToolTip = recentDocument.Path });
        if (recentItems.Count > 0) recentItems.AddRange([separator, new StudioMenuCommand("ClearRecent", Text("ClearRecent"), ViewModel.ClearRecentDocumentsCommand)]);

        var file = new List<StudioMenuEntry> {
            Command("Open"), new StudioMenuGroup("OpenRecent", Text("OpenRecent"), recentItems), separator,
            Command("Save"), Command("SaveAs"), separator,
            Command("Print"), separator,
            AsyncAction("CloseTab", CloseSelectedDocumentTabAsync, StudioShortcuts.CloseTab, () => TabHost.HasTabs),
            AsyncAction("ReopenTab", TabHost.ReopenClosedTabAsync, StudioShortcuts.ReopenTab, () => TabHost.CanReopenClosedTab)
        };
        if (!nativeMac) file.AddRange([separator, Command("Settings"), separator, Action("Exit", Close, StudioShortcuts.Quit)]);

        StudioMenuEntry[] edit = nativeMac
            ? [.. TextAwareEditEntries(), separator, Action("Find", FocusDocumentSearch, StudioShortcuts.Find), AsyncAction("Commands", ShowCommandPaletteAsync, StudioShortcuts.CommandSearch)]
            : [Command("Undo"), Command("Redo"), separator,
                Action("Find", FocusDocumentSearch, StudioShortcuts.Find, () => ViewModel.HasDocument),
                AsyncAction("Commands", ShowCommandPaletteAsync, StudioShortcuts.CommandSearch)];

        StudioMenuGroup view = Group("View",
            Command("Home"), Command("Tools"), Command("Jobs"), separator,
            Command("ZoomIn"), Command("ZoomOut"),
            Command("ActualSize"), Command("FitPage"), Command("FitWidth"), separator,
            Command("FocusReading"), Command("PageNightMode"), Command("CompareSideBySide"), Command("Assistant"),
            Action("ToggleSidebar", () => OnSidebarToggleClick(null, new()), StudioShortcuts.ToggleSidebar, () => ViewModel.HasDocument),
            Action(WindowState == WindowState.FullScreen ? "LeaveFullScreen" : "FullScreen", ToggleFullScreen, StudioShortcuts.FullScreen),
            separator, AppearanceMenu());

        StudioMenuGroup document = Group("Document",
            Command("Read"), Command("Comment"), Command("Edit"), Command("Pages"), Command("Forms"), Command("Protect"),
            separator, Command("Watermark"), Command("PageNumbers"), Command("FlattenAnnotations"),
            separator, Command("Split"),
            separator, Command("Redact"), Command("ValidateSignatures"),
            separator, Command("PreviousFormField"), Command("NextFormField"));

        // Tools follow the catalog's categories so the menu and the Tools page group them the same way.
        StudioMenuGroup tools = Group("ToolsMenu", commands.GroupTools(null)
            .Select(group => (StudioMenuEntry)new StudioMenuGroup(group.Title, group.Title, group.Items.Select(tool => (StudioMenuEntry)Command(tool.Id)).ToArray()))
            .ToArray());

        var window = new List<StudioMenuEntry> {
            Action("Minimize", () => WindowState = WindowState.Minimized, StudioShortcuts.Minimize),
            Action("NextTab", () => TabHost.SelectRelativeTab(false), StudioShortcuts.NextTab, () => TabHost.Tabs.Count > 1),
            Action("PreviousTab", () => TabHost.SelectRelativeTab(true), StudioShortcuts.PreviousTab, () => TabHost.Tabs.Count > 1)
        };
        if (TabHost.HasTabs) {
            window.Add(separator);
            foreach (StudioDocumentTabViewModel tab in TabHost.Tabs) {
                window.Add(new StudioMenuCommand("Tab", tab.DisplayTitle, new RelayCommand(() => { if (TabHost.Tabs.Contains(tab)) TabHost.SelectedTab = tab; })) {
                    IsChecked = ReferenceEquals(tab, TabHost.SelectedTab)
                });
            }
        }

        var help = new List<StudioMenuEntry> {
            AsyncAction("KeyboardShortcuts", ShowCommandPaletteAsync),
            AsyncAction("Documentation", () => OpenUriAsync(new Uri(DocumentationUri))),
            AsyncAction("ReportProblem", () => OpenUriAsync(new Uri(ReportProblemUri)))
        };
        if (!nativeMac) help.AddRange([separator, Action("About", ShowAbout)]);

        return [Group("File", [.. file]), Group("Edit", edit), view, document, tools, Group("Window", [.. window]), Group("Help", [.. help])];
    }

    // About opens Settings at the version and update details rather than wherever Settings was left.
    internal void ShowAbout() => ViewModel.ShowSettingsSection(Settings.StudioSettingsViewModel.AboutSection);

    private StudioMenuGroup AppearanceMenu() {
        string Text(string key) => _services.Localizer.Get("Settings.Theme" + key);
        StudioThemePreference current = _services.Preferences.Current.Theme;
        return new StudioMenuGroup("Appearance", _services.Localizer.Get("Settings.Appearance"), new[] {
                (StudioThemePreference.System, "System"), (StudioThemePreference.Light, "Light"),
                (StudioThemePreference.Dark, "Dark"), (StudioThemePreference.HighContrast, "HighContrast") }
            .Select(pair => (StudioMenuEntry)new StudioMenuCommand("Theme" + pair.Item2, Text(pair.Item2),
                new RelayCommand(() => _services.Preferences.UpdatePresentation(value => value with { Theme = pair.Item1 }))) {
                IsChecked = current == pair.Item1
            }).ToArray());
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

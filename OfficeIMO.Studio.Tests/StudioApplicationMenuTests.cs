using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioApplicationMenuTests {
    [Fact]
    public async Task Application_menu_is_built_from_the_command_catalog() {
        if (OperatingSystem.IsMacOS()) return;
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(() => {
            var services = TestAppBuilder.CreateTestServices();
            var window = new MainWindow(services) { Width = 1280, Height = 800 };
            try {
                window.Show();
                IReadOnlyList<MenuItem> menus = window.BuildAppMenu();
                Assert.Equal(new[] { "File", "Edit", "View", "Document", "ToolsMenu", "Window", "Help" }, menus.Select(menu => menu.Tag));
                Assert.All(menus, menu => Assert.StartsWith("_", (string)menu.Header!));

                // Every catalog tool appears once, grouped under its category.
                MenuItem tools = menus.Single(menu => Equals(menu.Tag, "ToolsMenu"));
                string[] listed = tools.Items.OfType<MenuItem>().SelectMany(category => category.Items.OfType<MenuItem>()).Select(item => (string)item.Tag!).ToArray();
                Assert.Equal(window.ViewModel.Commands.Items.Where(item => item.IsTool).Select(item => item.Id).Order(), listed.Order());

                // Unavailable commands are disabled and explain why.
                MenuItem file = menus.Single(menu => Equals(menu.Tag, "File"));
                MenuItem save = file.Items.OfType<MenuItem>().Single(item => Equals(item.Tag, "Save"));
                Assert.False(window.ViewModel.Commands["Save"].CanExecute(null));
                Assert.Equal(window.ViewModel.Commands["Save"].UnavailableReason, ToolTip.GetTip(save));
                Assert.Equal(KeyGesture.Parse("Ctrl+S"), save.InputGesture);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task F10_and_access_keys_open_the_application_menu() {
        if (OperatingSystem.IsMacOS()) return;
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(() => {
            var services = TestAppBuilder.CreateTestServices();
            var window = new MainWindow(services) { Width = 1280, Height = 800 };
            try {
                window.Show();
                Dispatcher.UIThread.RunJobs();
                window.KeyPress(Key.F10, RawInputModifiers.None, PhysicalKey.F10, null);
                Dispatcher.UIThread.RunJobs();
                Assert.True(window.AppMenu!.IsOpen);
                Assert.NotEmpty(window.AppMenu.Items);
                window.AppMenu.Hide();

                window.KeyPress(Key.H, RawInputModifiers.Alt, PhysicalKey.H, "h");
                Dispatcher.UIThread.RunJobs();
                Assert.True(window.AppMenu.IsOpen);
                MenuItem help = window.AppMenu.Items.OfType<MenuItem>().Single(item => Equals(item.Tag, "Help"));
                Assert.True(help.IsSubMenuOpen);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
}

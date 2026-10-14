using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Settings;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Diagnostics;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioOcrSetupTests {
    [Theory]
    [InlineData(960, 620)]
    [InlineData(1280, 800)]
    public async Task SetupIsReachableInRenderedSettingsAndCheckUpdatesItsResult(int width, int height) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                window.ViewModel.ShowSettingsCommand.Execute(null);
                window.Measure(new Size(width, height)); window.Arrange(new Rect(0, 0, width, height)); window.UpdateLayout();
                var view = window.GetVisualDescendants().OfType<OfficeIMO.Studio.Features.Settings.SettingsView>().Single(v => v.IsEffectivelyVisible);
                var setup = window.ViewModel.Settings.OcrSetup;
                var button = view.GetVisualDescendants().OfType<Button>().Single(b => ReferenceEquals(b.Command, setup.CheckCommand));
                button.BringIntoView(); window.UpdateLayout();
                Assert.True(button.IsEffectivelyVisible);
                Assert.True(button.Bounds.Width > 30);
                // Use an unavailable explicit path so the UI check remains deterministic and downloads nothing.
                setup.ExecutablePath = Path.Combine(services.Paths.Root, "missing-tesseract");
                await setup.CheckCommand.ExecuteAsync(null);
                Assert.False(setup.IsReady);
                Assert.Null(services.Preferences.Current.OcrExecutablePath);
                Assert.Contains("not found", setup.Status);
                Assert.True(setup.CanCheck);
                Assert.True(setup.CheckCommand.CanExecute(null));
                await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => { }, Avalonia.Threading.DispatcherPriority.Background);
                window.UpdateLayout();
                Assert.True(button.IsEnabled);
                var scroll = view.GetVisualDescendants().OfType<ScrollViewer>().First();
                scroll.Offset = new Vector(0, 370);
                window.UpdateLayout();
                using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
                string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrWhiteSpace(output)) {
                    Directory.CreateDirectory(output);
                    frame.Save(Path.Combine(output, $"ocr-setup-{width}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task SharedSettingsRefreshSavedChoiceWithoutOverwritingAnotherDraft() {
        var preferences = new StudioPreferencesService(new MemoryPreferences());
        using var first = new StudioOcrSetupViewModel(preferences, English(), new NoDiagnostics(),
            inspect: (path, _) => Task.FromResult(new StudioOcrRuntimeInfo(path!, "tesseract 5", ["eng"])));
        using var other = new StudioOcrSetupViewModel(preferences, English(), new NoDiagnostics());
        first.ExecutablePath = Path.GetFullPath("engine-one");
        await first.CheckCommand.ExecuteAsync(null);
        Assert.Equal(first.ExecutablePath, other.ExecutablePath);
        other.ExecutablePath = Path.GetFullPath("draft-engine");
        first.ExecutablePath = Path.GetFullPath("engine-two");
        await first.CheckCommand.ExecuteAsync(null);
        Assert.Equal(Path.GetFullPath("draft-engine"), other.ExecutablePath);
        Assert.Equal(first.ExecutablePath, preferences.Current.OcrExecutablePath);
    }

    [Fact]
    public async Task VerifiedExecutableIsPersistedAndFailedReplacementPreservesIt() {
        var store = new MemoryPreferences();
        var preferences = new StudioPreferencesService(store);
        bool fail = false;
        string executable = Path.GetFullPath("verified-tesseract");
        using var model = new StudioOcrSetupViewModel(preferences, English(), new NoDiagnostics(),
            inspect: (path, _) => fail ? throw new FileNotFoundException() :
                Task.FromResult(new StudioOcrRuntimeInfo(path!, "tesseract 5.5.1", ["eng", "pol"])));
        model.ExecutablePath = executable;
        await model.CheckCommand.ExecuteAsync(null);
        Assert.True(model.IsReady);
        Assert.Equal(executable, new StudioPreferencesService(store).Current.OcrExecutablePath);
        Assert.Equal("eng, pol", model.InstalledLanguages);
        fail = true;
        model.ExecutablePath = Path.GetFullPath("missing-engine");
        await model.CheckCommand.ExecuteAsync(null);
        Assert.False(model.IsReady);
        Assert.Null(model.Version);
        Assert.Equal(executable, preferences.Current.OcrExecutablePath);
        Assert.Contains("not found", model.Status);
    }

    [Fact]
    public async Task AutomaticDiscoveryClearsExplicitPreferenceOnlyAfterSuccessfulCheck() {
        var store = new MemoryPreferences();
        var preferences = new StudioPreferencesService(store);
        preferences.Update(p => p with { OcrExecutablePath = Path.GetFullPath("old-engine") });
        bool fail = true;
        using var model = new StudioOcrSetupViewModel(preferences, English(), new NoDiagnostics(),
            inspect: (path, _) => {
                Assert.Null(path);
                if (fail) throw new FileNotFoundException();
                return Task.FromResult(new StudioOcrRuntimeInfo(Path.GetFullPath("auto-engine"), "tesseract 5", []));
            });
        await model.UseAutomaticDiscoveryCommand.ExecuteAsync(null);
        Assert.NotNull(preferences.Current.OcrExecutablePath);
        fail = false;
        await model.UseAutomaticDiscoveryCommand.ExecuteAsync(null);
        Assert.Null(preferences.Current.OcrExecutablePath);
        Assert.True(model.IsReady);
        Assert.Contains("No recognition languages", model.InstalledLanguages);
    }

    [Fact]
    public async Task RelativeExecutableIsRejectedWithoutInvokingAProcess() {
        using var model = new StudioOcrSetupViewModel(new(new MemoryPreferences()), English(), new NoDiagnostics(),
            inspect: (_, _) => throw new Xunit.Sdk.XunitException("An ambiguous executable must never be probed."));
        model.ExecutablePath = "tesseract";
        await model.CheckCommand.ExecuteAsync(null);
        Assert.False(model.IsReady);
        Assert.Contains("complete file path", model.Status);
    }

    [Fact]
    public async Task ClosingSetupCancelsProbeAndDoesNotPersistItsResult() {
        var preferences = new StudioPreferencesService(new MemoryPreferences());
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using var model = new StudioOcrSetupViewModel(preferences, English(), new NoDiagnostics(),
            inspect: async (_, token) => {
                entered.SetResult();
                await Task.Delay(Timeout.Infinite, token);
                throw new InvalidOperationException("Unreachable");
            });
        var check = model.CheckCommand.ExecuteAsync(null);
        await entered.Task.WaitAsync(TimeSpan.FromSeconds(5));
        model.Dispose();
        await check.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.Null(preferences.Current.OcrExecutablePath);
    }

    private static StudioLocalizer English() => new(CultureInfo.GetCultureInfo("en"));
    private sealed class MemoryPreferences : IStudioPreferencesStore {
        private StudioPreferences _value = new();
        public StudioPreferences Load() => _value;
        public void Save(StudioPreferences preferences) => _value = preferences;
    }
    private sealed class NoDiagnostics : IStudioDiagnostics {
        public string DirectoryPath => string.Empty;
        public void Write(StudioDiagnosticLevel level, string area, string code, Exception? exception = null) { }
        public StudioSupportSnapshot CreateSupportSnapshot() => throw new NotSupportedException();
    }
}

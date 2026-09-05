using Avalonia;
using Avalonia.Controls;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioViewPreferencesTests {
    [Fact]
    public void StoredViewPreferencesAreBoundedAndDoNotContainSourcePaths() {
        string root = NewFolder();
        try {
            string file = Path.Combine(root, "views.json");
            var store = new StudioDocumentViewStore(file);
            string document = Path.Combine(root, "private-source.pdf");
            store.Put(document, new StudioDocumentViewState {
                NavigationWidth = double.NaN, InspectorWidth = 9000, PageNumber = -20, Zoom = 20,
                Panes = new() { [StudioDocumentMode.View] = new(false, false), [(StudioDocumentMode)99] = new(true, true) }
            });
            var loaded = new StudioDocumentViewStore(file).Get(document);
            Assert.Equal(238, loaded.NavigationWidth);
            Assert.Equal(380, loaded.InspectorWidth);
            Assert.Equal(1, loaded.PageNumber);
            Assert.Equal(3, loaded.Zoom);
            Assert.Single(loaded.Panes);
            Assert.DoesNotContain("private-source", File.ReadAllText(file));
            for (int i = 0; i < 65; i++) store.Put(Path.Combine(root, i + ".pdf"), new() { PageNumber = 2 });
            Assert.Equal(1, new StudioDocumentViewStore(file).Get(document).PageNumber);
            Assert.Equal(2, new StudioDocumentViewStore(file).Get(Path.Combine(root, "64.pdf")).PageNumber);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task ReopeningRestoresReadingPositionZoomAndPerDocumentPanePreferences() {
        string root = NewFolder();
        try {
            string source = Path.Combine(root, "reading.pdf");
            PdfDocument.Create(compose => {
                compose.Page(page => page.Size(600, 800));
                compose.Page(page => page.Size(600, 800));
            }).Save(source);
            var paths = new StudioDataPaths(Path.Combine(root, "profile"));
            using (var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: StudioApplicationServices.Create(paths))) {
                await model.OpenDocumentAsync(source);
                model.SelectedPage = model.Pages[1];
                model.ActualSizeCommand.Execute(null);
                model.ZoomInCommand.Execute(null);
                model.UpdatePanePreferences(310, 370, new(false, true));
                model.ToggleFocusReadingCommand.Execute(null);
            }
            using var reopened = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: StudioApplicationServices.Create(paths));
            await reopened.OpenDocumentAsync(source);
            Assert.Equal(2, reopened.SelectedPage!.PageNumber);
            Assert.Equal(1.25, reopened.Zoom);
            Assert.Equal(310, reopened.DocumentViewState.NavigationWidth);
            Assert.Equal(370, reopened.DocumentViewState.InspectorWidth);
            Assert.Equal(new StudioPanePreference(false, true), reopened.DocumentViewState.Panes[StudioDocumentMode.View]);
            Assert.True(reopened.IsFocusReading);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task DensityChangesControlSizesLiveAndFocusReadingRestoresPanes() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var app = (App)Application.Current!;
            var window = new MainWindow();
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                var workspace = window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!;
                var button = workspace.FitPageButtonControl;
                Assert.Equal(38, button.MinHeight);
                app.Services.Preferences.Update(current => current with { Density = StudioDensityPreference.Compact });
                window.UpdateLayout();
                Assert.Equal(32, button.MinHeight);
                app.Services.Preferences.Update(current => current with { Density = StudioDensityPreference.Comfortable });
                window.UpdateLayout();
                Assert.Equal(38, button.MinHeight);
                window.ViewModel.ShowAnnotateModeCommand.Execute(null);
                Assert.True(workspace.FindControl<Grid>("InspectorPane")!.IsVisible);
                window.ViewModel.ToggleFocusReadingCommand.Execute(null);
                Assert.False(workspace.FindControl<Grid>("InspectorPane")!.IsVisible);
                Assert.False(workspace.FindControl<Grid>("NavigationPane")!.IsVisible);
                window.ViewModel.ToggleFocusReadingCommand.Execute(null);
                Assert.Equal(StudioDocumentMode.View, window.ViewModel.DocumentMode);
                Assert.True(workspace.FindControl<Grid>("NavigationPane")!.IsVisible);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static string NewFolder() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-view-preferences-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }
}

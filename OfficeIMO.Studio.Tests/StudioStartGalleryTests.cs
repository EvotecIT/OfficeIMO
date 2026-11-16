using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Home;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioStartGalleryTests {
    [Fact]
    public async Task GalleryShowsEveryToolOnceAndOnlyGuardedCommandsAskForADocument() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            string[] shown = model.Commands.Start.Groups.SelectMany(group => group.Tiles).Select(tile => tile.Id).ToArray();
            Assert.Equal(shown.Length, shown.Distinct().Count());
            Assert.All(model.Commands.Start.Groups, group => Assert.NotEmpty(group.Tiles));
            // Every catalog tool and every window-level workbench has a place on the start screen.
            Assert.Empty(model.Commands.Items.Where(item => item.IsTool).Select(item => item.Id).Except(shown));
            Assert.Contains("Jobs", shown);

            string openFirst = model.Commands["Edit"].UnavailableReason!;
            foreach (StudioStartTile tile in model.Commands.Start.Groups.SelectMany(group => group.Tiles)) {
                Assert.True(tile.CanExecute(null), tile.Id);
                Assert.Equal(tile.Command.RequiresDocument, tile.OpensDocumentFirst);
                if (tile.OpensDocumentFirst) Assert.Equal(openFirst, tile.Command.UnavailableReason);
                else Assert.True(tile.Command.IsAvailable, tile.Id);
            }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task DocumentToolAsksForAPdfAndRunsOnTheOpenedDocument() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-start-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "start.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(500))
            .Content(content => content.AddParagraph("Start screen")))).Save(path);
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                string? choice = null;
                using var model = new MainWindowViewModel(_ => Task.FromResult(choice));
                await model.StartToolAsync("Pages");
                Assert.False(model.HasDocument);
                Assert.Equal(StudioWorkspaceMode.Home, model.WorkspaceMode);

                choice = path;
                await model.StartToolAsync("Pages");
                Assert.True(model.HasDocument, model.ErrorMessage);
                Assert.Equal(StudioWorkspaceMode.PdfWorkspace, model.WorkspaceMode);
                Assert.Equal(StudioDocumentMode.Pages, model.DocumentMode);
                Assert.False(model.Commands.Start["Pages"].OpensDocumentFirst);

                // From the start screen with a document open, a tool that doesn't switch screens itself still shows the document.
                model.WorkspaceMode = StudioWorkspaceMode.Home;
                bool nightMode = model.IsPageNightMode;
                await model.StartToolAsync("PageNightMode");
                Assert.Equal(StudioWorkspaceMode.PdfWorkspace, model.WorkspaceMode);
                Assert.NotEqual(nightMode, model.IsPageNightMode);

                // Validation results live in the Protect panel's certificate signature section, which opens with them.
                model.WorkspaceMode = StudioWorkspaceMode.Home;
                await model.StartToolAsync("ValidateSignatures");
                Assert.Equal(StudioWorkspaceMode.PdfWorkspace, model.WorkspaceMode);
                Assert.Equal(StudioDocumentMode.Protect, model.DocumentMode);
                Assert.True(model.IsCertificateSignatureSectionExpanded);
                Assert.True(model.HasSignatureValidationSummary);

                // Workbenches keep their own screen.
                model.WorkspaceMode = StudioWorkspaceMode.Home;
                await model.StartToolAsync("Invoices");
                Assert.True(model.IsInvoiceMode);
                return true;
            }, CancellationToken.None);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData(1280, 2600, false)]
    [InlineData(1280, 2600, true)]
    [InlineData(520, 2400, false)]
    [InlineData(520, 2400, true)]
    public async Task ColdStartShowsTheStartScreenAndATileOpensItsWorkbench(int width, int height, bool dark) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(() => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            var recent = new List<RecentDocumentViewModel>();
            foreach (string name in new[] { "Quarterly report.pdf", "Signed contract.pdf" }) {
                string path = Path.Combine(services.Paths.Root, name);
                PdfDocument.Create(document => document.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(500))
                    .Content(content => content.AddParagraph(name)))).Save(path);
                recent.Add(new RecentDocumentViewModel(path, DateTimeOffset.Now));
            }
            services.DocumentHistory.RecentDocuments.Save(recent);
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                // A cold start with no arguments or session shows the start screen, with the stored recent documents.
                Assert.True(window.ViewModel.IsHomeMode);
                Assert.Equal(2, window.ViewModel.RecentDocuments.Count);
                Layout(window, width, height);
                var home = window.GetVisualDescendants().OfType<HomeView>().Single(view => view.IsEffectivelyVisible);
                Button[] tiles = home.GetVisualDescendants().OfType<Button>().Where(button => button.Classes.Contains("galleryTile")).ToArray();
                Assert.Equal(window.ViewModel.Commands.Start.Groups.Sum(group => group.Tiles.Count), tiles.Length);
                Assert.All(tiles, tile => {
                    Assert.True(tile.IsEffectivelyEnabled);
                    Assert.False(string.IsNullOrWhiteSpace(Avalonia.Automation.AutomationProperties.GetName(tile)));
                    Point origin = tile.TranslatePoint(default, home)!.Value;
                    Assert.InRange(origin.X + tile.Bounds.Width, 0, home.Bounds.Width + 1);
                });
                Capture(window, width, dark, "start");

                Button invoices = tiles.Single(tile => tile.DataContext is StudioStartTile { Id: "Invoices" });
                invoices.Command!.Execute(invoices.CommandParameter);
                Assert.True(window.ViewModel.IsInvoiceMode);
                Layout(window, width, Math.Min(height, 900));
                Capture(window, width, dark, "invoices");
            } finally {
                window.Close();
            }
            return true;
        }, CancellationToken.None);
    }

    private static void Layout(MainWindow window, int width, int height) {
        window.Width = width;
        window.Height = height;
        window.ApplyResponsiveLayout(width);
        window.Measure(new Size(width, height));
        window.Arrange(new Rect(0, 0, width, height));
        window.UpdateLayout();
    }

    private static void Capture(Window window, int width, bool dark, string state) {
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, $"start-{width}-{(dark ? "dark" : "light")}-{state}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}

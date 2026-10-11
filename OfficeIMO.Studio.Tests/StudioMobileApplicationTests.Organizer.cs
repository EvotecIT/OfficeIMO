using Avalonia;
using Avalonia.Controls;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed partial class StudioMobileApplicationTests {
    [Fact]
    public async Task OrganizerActionsRemainClickableWhenIndependentReadersAreHidden() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-organizer-readers-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "three-pages.pdf");
        PdfDocument.Create(pdf => {
            for (int page = 1; page <= 3; page++)
                pdf.Page(item => item.Content(content => content.Text($"Page {page}")));
        }).Save(source);
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = TestAppBuilder.CreateTestServices();
                var view = new DocumentWorkspaceView();
                using var host = new StudioDocumentTabHost(open => new MainWindowViewModel(
                    _ => Task.FromResult<string?>(null), openDocumentInTab: open, services: services),
                    document => view.DataContext = document);
                view.DataContext = host.ActiveDocument;
                host.ActiveDocument.ShowPdfWorkspaceCommand.Execute(null);
                view.Panes = host.Panes;
                var window = new Window { Content = view, Width = 1366, Height = 1024 };
                try {
                    window.Show(); Layout(window);
                    await ClickAsync(window, view, host.ActiveDocument.Commands["Open"]);
                    await host.ActiveDocument.OpenCommand.ExecutionTask!;
                    Assert.False(host.ActiveDocument.HasDocument);
                    Capture(window, "organizer-empty-host");
                    await host.OpenDocumentAsync(source);
                    var document = host.ActiveDocument;
                    host.Panes.OpenSecondPane(host.SelectedTab);
                    await document.Commands["Pages"].ExecuteAsync(); Layout(window);
                    document.OrganizerPageRange = "2";
                    document.SelectPageRangeCommand.Execute(null);
                    Assert.True(document.HasOrganizerSelection);
                    using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                    while (document.OrganizerPages.Any(page => page.IsLoading)) await Task.Delay(10, timeout.Token);
                    Layout(window); Capture(window, "organizer-independent-pages");
                    await ClickAsync(window, view, document.MoveSelectedUpCommand);
                    await document.MoveSelectedUpCommand.ExecutionTask!;
                    Assert.True(document.IsDirty);
                    await document.SaveCommand.ExecuteAsync(null);
                    Assert.False(document.IsDirty, document.ErrorMessage);
                    Assert.Equal(new[] { "Page 2", "Page 1", "Page 3" },
                        PdfReadDocument.Open(document.DocumentPath!).Pages.Select(page => page.ExtractText().Trim()));
                    document.ShowViewModeCommand.Execute(null); Layout(window);
                    foreach (var pane in new[] { host.Panes.Left!, host.Panes.Right! }) {
                        await pane.SelectedPage!.EnsureRenderedAsync();
                        Assert.True(pane.SelectedPage.HasScene);
                    }
                    Layout(window); Capture(window, "organizer-independent-readers-return");
                } finally { window.Close(); services.Storage.Dispose(); }
                return true;
            }, CancellationToken.None);
        } finally { Directory.Delete(root, recursive: true); }
    }
}

using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioWorkflowFeedbackTests {
    [Theory]
    [InlineData(960)]
    [InlineData(1600)]
    public async Task FormsToolReopensPreviouslyHiddenProperties(int width) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Avalonia.Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "forms-navigation.pdf");
            PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("Forms")))).Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await model.OpenDocumentAsync(source);
            model.DocumentMode = StudioDocumentMode.Forms;
            model.UpdatePanePreferences(238, 300, new(false, false));
            var view = new Features.Workspace.DocumentWorkspaceView { DataContext = model };
            var window = new Avalonia.Controls.Window { Content = view, Width = width, Height = 720 };
            try {
                window.Show();
                view.ApplyResponsiveLayout(width);
                var inspector = Avalonia.Controls.ControlExtensions.FindControl<Avalonia.Controls.Control>(view, "InspectorPane")!;
                Assert.False(inspector.IsVisible);
                model.ShowToolsCommand.Execute(null);
                await model.Commands["Forms"].ExecuteAsync();
                Assert.Equal(StudioWorkspaceMode.PdfWorkspace, model.WorkspaceMode);
                Assert.Equal(StudioDocumentMode.Forms, model.DocumentMode);
                Assert.True(inspector.IsVisible);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RealOperationCompletionRetainsOriginAfterNavigation(bool signatures) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Avalonia.Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "notification-source.pdf");
            string output = Path.Combine(services.Paths.Root, "notification-copy.pdf");
            PdfDocument.Create(document => document.Page(page => page.Content(content =>
                content.Text("Operation notification proof")))).Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(output), reviewPageExtraction: _ => Task.FromResult(true));
            await model.OpenDocumentAsync(source);
            model.WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
            model.DocumentMode = signatures ? StudioDocumentMode.Protect : StudioDocumentMode.Pages;
            model.SelectAllPagesCommand.Execute(null);
            using var blocker = await Features.Workspace.PdfWorkspace.OpenAsync(source, CancellationToken.None);
            var acquired = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            Task<bool> holding = blocker.RunNonDetachableCpuWorkAsync(() => {
                acquired.SetResult(); release.Task.GetAwaiter().GetResult(); return true;
            }, CancellationToken.None);
            try {
                await acquired.Task;
                Task running = signatures ? model.ValidateSignaturesCommand.ExecuteAsync(null)
                    : model.ExtractSelectedCommand.ExecuteAsync(null);
                Assert.True(model.IsWorkspaceBusy);
                model.WorkspaceMode = StudioWorkspaceMode.Tools;
                release.SetResult();
                await running;
                Assert.False(model.HasVisibleOperationStatus);
                Assert.False(model.HasVisibleError);
                Assert.False(model.HasError);
                if (signatures) Assert.Equal("Signature validation complete", model.OperationStatus);
                else Assert.True(File.Exists(output));
                model.WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
                Assert.True(model.HasVisibleOperationStatus);
                model.WorkspaceMode = StudioWorkspaceMode.Tools;
                model.OperationStatus = "Choose a tool";
                Assert.True(model.HasVisibleOperationStatus);
            } finally { release.TrySetResult(); await holding; }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task DocumentFailureStaysWithItsTaskAndDoesNotDuplicateTheStatusBanner() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            model.WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
            model.DocumentMode = StudioDocumentMode.Edit;
            model.ErrorMessage = "The edit could not be applied.";
            model.OperationStatus = "Operation failed";
            Assert.True(model.HasVisibleError);
            Assert.False(model.HasVisibleOperationStatus);
            foreach (var workspace in new[] { StudioWorkspaceMode.Home, StudioWorkspaceMode.Tools,
                StudioWorkspaceMode.Ocr, StudioWorkspaceMode.Convert, StudioWorkspaceMode.Settings }) {
                model.WorkspaceMode = workspace;
                Assert.False(model.HasVisibleError);
                Assert.False(model.HasVisibleOperationStatus);
            }
            model.WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
            model.DocumentMode = StudioDocumentMode.Forms;
            Assert.False(model.HasVisibleError);
            model.DocumentMode = StudioDocumentMode.Edit;
            Assert.True(model.HasVisibleError);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task LateCompletionKeepsTheOriginWhenTheUserChangesWorkspace() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            model.WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
            model.DocumentMode = StudioDocumentMode.Edit;
            model.IsWorkspaceBusy = true;
            model.WorkspaceMode = StudioWorkspaceMode.Tools;
            model.OperationStatus = "Selected object moved.";
            model.IsWorkspaceBusy = false;
            Assert.False(model.HasVisibleOperationStatus);
            model.WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
            Assert.True(model.HasVisibleOperationStatus);
            model.IsWorkspaceBusy = true;
            model.WorkspaceMode = StudioWorkspaceMode.Tools;
            model.OperationStatus = "Operation completed";
            model.IsWorkspaceBusy = false;
            model.OperationStatus = "Choose a tool";
            Assert.True(model.HasVisibleOperationStatus);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task MismatchedPdfRemainsAvailableForAnExplicitCompatibleRoute() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            string path = Path.Combine(Path.GetTempPath(), "selected-invoice.pdf");
            using var model = new ConversionWorkbenchViewModel(_ => Task.FromResult<IReadOnlyList<string>>([path]),
                _ => Task.FromResult<string?>(null));
            Assert.Equal("docx-pdf", model.SelectedRoute.Route.Id);
            await model.AddFilesCommand.ExecuteAsync(null);
            Assert.Empty(model.Jobs);
            Assert.True(model.HasUnmatchedInputs);
            Assert.All(model.MatchingInputRoutes, route => Assert.Contains(".pdf", route.Route.SourceExtensions));
            var selected = Assert.Single(model.MatchingInputRoutes, route => route.Route.Id == "pdf-html");
            model.UseInputRouteCommand.Execute(selected);
            var job = Assert.Single(model.Jobs);
            Assert.Equal(path, job.InputPath);
            Assert.Equal("pdf-html", job.Route.Route.Id);
            Assert.False(model.HasUnmatchedInputs);
            return true;
        }, CancellationToken.None);
    }
}

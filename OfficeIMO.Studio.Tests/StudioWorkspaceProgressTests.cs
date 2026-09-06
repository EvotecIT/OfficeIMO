using Avalonia.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioWorkspaceProgressTests {
    [Theory]
    [InlineData("Needle", 1)]
    [InlineData("Absent", 0)]
    public async Task SearchKeepsItsResultSummaryAfterTheOperationFinishes(string query, int count) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "search.pdf");
            PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("Needle in a document")))).Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await model.OpenDocumentAsync(source);
            model.SearchQuery = query;
            await model.SearchCommand.ExecuteAsync(null);
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            Assert.Equal(count, model.SearchResults.Count);
            Assert.Equal(count == 0 ? services.Localizer.Get("Workspace.NoMatches") : services.Localizer.Format("Workspace.MatchingPages", count), model.OperationStatus);
            Assert.Equal(1D, model.OperationProgressFraction);
            return true;
        }, CancellationToken.None);
    }
}

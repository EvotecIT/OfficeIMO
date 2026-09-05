using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioCommandCatalogTests {
    [Fact]
    public async Task UnavailableDocumentToolExplainsItsRequirementAndDoesNotChangeMode() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            model.ShowToolsCommand.Execute(null);
            var edit = model.Commands["Edit"];
            Assert.False(edit.IsAvailable);
            Assert.False(string.IsNullOrWhiteSpace(edit.UnavailableReason));
            await edit.ExecuteAsync();
            Assert.Equal(StudioWorkspaceMode.Tools, model.WorkspaceMode);
            Assert.True(model.Commands["Convert"].IsAvailable);
            await model.Commands["Convert"].ExecuteAsync();
            Assert.Equal(StudioWorkspaceMode.Convert, model.WorkspaceMode);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task CatalogObservesDocumentStateAndRunsTheExistingUndoTransaction() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-command-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "editable.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(600, 800)
            .Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Command transaction")))))).Save(path);
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
                var catalog = model.Commands;
                Assert.False(catalog["Read"].IsAvailable);
                await model.OpenDocumentAsync(path);
                Assert.True(catalog["Read"].IsAvailable);
                Assert.False(catalog["Undo"].IsAvailable);
                int pages = model.Pages.Count;
                await model.InsertBlankCommand.ExecuteAsync(null);
                Assert.False(model.HasError, model.ErrorMessage);
                Assert.Equal(pages + 1, model.Pages.Count);
                Assert.True(catalog["Undo"].IsAvailable);
                Assert.True(catalog["Save"].IsAvailable);
                await catalog["Undo"].ExecuteAsync();
                Assert.Equal(pages, model.Pages.Count);
                Assert.True(catalog["Redo"].IsAvailable);
                model.ShowToolsCommand.Execute(null);
                await catalog["Read"].ExecuteAsync();
                Assert.Equal(StudioWorkspaceMode.PdfWorkspace, model.WorkspaceMode);
                Assert.Equal(StudioDocumentMode.View, model.DocumentMode);
                return true;
            }, CancellationToken.None);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task SearchUsesDescriptionAndCategoryAndRetainsUnavailableChoices() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            model.Commands.ToolQuery = "scanned";
            Assert.Equal("Ocr", Assert.Single(model.Commands.FilteredTools).Id);
            model.Commands.ToolQuery = "OCR";
            Assert.Equal("Ocr", Assert.Single(model.Commands.FilteredTools).Id);
            model.Commands.ToolQuery = "redact";
            Assert.False(Assert.Single(model.Commands.FilteredTools).IsAvailable);
            var palette = new StudioCommandPaletteModel(model.Commands) { Query = "nonexistent-word" };
            Assert.False(palette.HasResults);
            Assert.Null(palette.SelectedCommand);
            palette.Query = "save copy";
            Assert.Equal("SaveAs", Assert.Single(palette.Results).Id);
            Assert.Same(palette.Results[0], palette.SelectedCommand);
            return true;
        }, CancellationToken.None);
    }
}

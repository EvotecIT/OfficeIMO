using Avalonia.Input;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioEditorToolTests {
    [Fact]
    public void Each_tool_key_is_unique_within_its_mode() {
        foreach (var mode in StudioEditorTools.All.GroupBy(tool => tool.Mode)) {
            var keys = mode.Where(tool => tool.Key is not null).Select(tool => tool.Key).ToArray();
            Assert.Equal(keys.Length, keys.Distinct().Count());
        }
        Assert.Equal("Highlight", StudioEditorTools.ForKey(StudioDocumentMode.Annotate, Key.H)!.Tool);
        Assert.Equal("Link", StudioEditorTools.ForKey(StudioDocumentMode.Edit, Key.K)!.Tool);
        Assert.Null(StudioEditorTools.ForKey(StudioDocumentMode.View, Key.H));
    }

    [Fact]
    public async Task Choosing_a_tool_from_search_switches_to_its_mode() {
        string fixture = Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf");
        using var document = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        await document.OpenDocumentAsync(fixture);

        StudioCommandItem highlight = Assert.Single(document.Commands.Search("highlight"), item => item.Id == "Tool.Highlight");
        Assert.Equal("H", highlight.Shortcut);
        Assert.True(highlight.CanExecute(null));

        await highlight.ExecuteAsync();

        Assert.True(document.IsAnnotateDocumentMode);
        Assert.Equal(PdfEditorTool.Highlight, document.ActiveEditorTool);
        Assert.All(StudioEditorTools.Searchable, tool => Assert.NotNull(document.Commands[tool.CommandId]));
    }
}

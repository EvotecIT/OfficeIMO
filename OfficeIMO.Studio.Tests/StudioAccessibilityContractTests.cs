using Avalonia.Automation.Peers;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Headless;
using Avalonia.LogicalTree;
using Avalonia.VisualTree;
using PdfDocument = OfficeIMO.Pdf.PdfDocument;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioAccessibilityContractTests {
    [Fact]
    public async Task RedactionMarksAnnounceMatchedContentAndTheirPage() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-redaction-accessibility-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            PdfDocument.Create(compose => {
                for (int page = 0; page < 2; page++) {
                    compose.Page(body => body.Content(content => {
                        content.Item(item => item.Paragraph(text => text.Text("Private account one")));
                        content.Item(item => item.Paragraph(text => text.Text("Private account two")));
                    }));
                }
            }).Save(source);
            using HeadlessUnitTestSession session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                MainWindow window = new MainWindow { Width = 1280, Height = 900 };
                try {
                    window.Show();
                    await window.TabHost.OpenDocumentAsync(source);
                    window.ViewModel.ShowProtectModeCommand.Execute(null);
                    window.ViewModel.RedactionSearchText = "Private account (one|two)";
                    window.ViewModel.RedactionSearchRegex = true;
                    await window.ViewModel.SearchRedactionsCommand.ExecuteAsync(null);
                    window.UpdateLayout();
                    Expander[] sections = window.GetVisualDescendants().OfType<Expander>()
                        .Where(section => section.Classes.Contains("accessibleSection")).ToArray();
                    Assert.Equal(4, sections.Length);
                    foreach (Expander section in sections) {
                        TextBlock title = Assert.Single(Assert.IsType<StackPanel>(section.Header).Children.OfType<TextBlock>());
                        ToggleButton header = Assert.Single(section.GetVisualDescendants().OfType<ToggleButton>(),
                            button => button.TemplatedParent == section && button.Name == "ExpanderHeader");
                        Assert.Equal(title.Text, ControlAutomationPeer.CreatePeerForElement(header)!.GetName());
                    }
                    RedactionInspectorView inspector = Assert.Single(window.GetVisualDescendants().OfType<RedactionInspectorView>());
                    ListBox marks = Assert.Single(inspector.GetVisualDescendants().OfType<ListBox>());
                    Assert.Equal(4, window.ViewModel.RedactionMarks.Count);
                    var includeNames = new List<string>();
                    for (int index = 0; index < 4; index++) {
                        ListBoxItem row = Assert.IsType<ListBoxItem>(marks.ContainerFromIndex(index));
                        string name = ControlAutomationPeer.CreatePeerForElement(row)!.GetName();
                        Assert.StartsWith($"Page {index / 2 + 1}:", name, StringComparison.Ordinal);
                        Assert.Contains("Private account", name, StringComparison.Ordinal);
                        CheckBox include = Assert.Single(row.GetVisualDescendants().OfType<CheckBox>());
                        string includeName = ControlAutomationPeer.CreatePeerForElement(include)!.GetName();
                        Assert.Equal($"Include page {index / 2 + 1}: {window.ViewModel.RedactionMarks[index].Description}", includeName);
                        includeNames.Add(includeName);
                    }
                    Assert.Equal(4, includeNames.Distinct(StringComparer.Ordinal).Count());
                } finally {
                    foreach (StudioDocumentTabViewModel tab in window.TabHost.Tabs.ToArray()) tab.Document.CompletePreparedClose();
                    window.Close();
                }
                return true;
            }, CancellationToken.None);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task PdfPageCanvasExposesDocumentTextThroughAutomationTree() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            using var canvas = new PdfPageCanvas { Scene = TestPdfPageScenes.Create() };
            var peer = new PdfPageCanvasAutomationPeer(canvas);

            Assert.Equal(AutomationControlType.Document, peer.GetAutomationControlType());
            AutomationPeer text = Assert.Single(
                peer.GetChildren(),
                child => child.GetAutomationControlType() == AutomationControlType.Text);
            Assert.Contains("Page text", text.GetName(), StringComparison.OrdinalIgnoreCase);
            Assert.False(string.IsNullOrWhiteSpace(text.GetName()));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task UnsavedChangesDialogProvidesDefaultAndCancelTargets() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var unsaved = new UnsavedChangesDialog("sample.pdf");
            var unsavedButtons = GetButtons(unsaved);
            Assert.Contains(unsavedButtons, button => button.IsDefault);
            Assert.Contains(unsavedButtons, button => button.IsCancel);

            return true;
        }, CancellationToken.None);
    }

    private static Button[] GetButtons(Window window) {
        return window.GetLogicalDescendants().OfType<Button>().ToArray();
    }
}

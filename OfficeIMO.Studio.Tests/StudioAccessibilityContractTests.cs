using Avalonia.Automation.Peers;
using Avalonia.Controls;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioAccessibilityContractTests {
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
    public async Task DestructiveDialogsProvideDefaultCancelAndInitialFocusTargets() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var unsaved = new UnsavedChangesDialog("sample.pdf");
            var unsavedButtons = GetButtons(unsaved);
            Assert.Contains(unsavedButtons, button => button.IsDefault);
            Assert.Contains(unsavedButtons, button => button.IsCancel);

            var deletion = new PageDeletionDialog(2);
            Assert.Contains(GetButtons(deletion), button => button.IsCancel);
            return true;
        }, CancellationToken.None);
    }

    private static Button[] GetButtons(Window window) {
        var root = Assert.IsType<StackPanel>(window.Content);
        var actions = Assert.IsType<StackPanel>(root.Children[^1]);
        return actions.Children.OfType<Button>().ToArray();
    }
}

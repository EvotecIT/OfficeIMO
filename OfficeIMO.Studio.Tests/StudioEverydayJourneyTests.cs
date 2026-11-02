using Avalonia.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

/// <summary>
/// The consistent-workspace journey end to end: open a document, find text, change tasks, edit an object, undo and
/// redo, and save, keeping the reader's place throughout.
/// </summary>
public sealed class StudioEverydayJourneyTests {
    [Fact]
    public async Task Find_edit_undo_and_save_keep_the_readers_place() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "journey.pdf");
            string saved = Path.Combine(services.Paths.Root, "journey-reviewed.pdf");
            PdfDocument.Create(document => {
                foreach (string text in new[] { "Introduction", "Quarterly results", "Appendix" }) {
                    document.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(560)).Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text(text)))));
                }
            }).Save(source);
            using var document = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(saved));
            await document.OpenDocumentAsync(source);

            // Find text: the result's page becomes the reader's place.
            document.ZoomInCommand.Execute(null);
            double zoom = document.Zoom;
            document.SearchQuery = "Quarterly";
            await document.SearchCommand.ExecuteAsync(null);
            Assert.True(document.HasSearchResults);
            Assert.Equal(2, document.SelectedPage!.PageNumber);

            // Change task and add a note on that page.
            document.ShowAnnotateModeCommand.Execute(null);
            document.SelectedEditorToolChoice = document.EditorTools.Single(choice => choice.Tool == PdfEditorTool.Note);
            document.Pages[1].CompleteEditorGesture(new PdfEditorGesture(2, 60, 60, 78, 78, [new PdfEditorVisualPoint(60, 60)]));
            // The edit enters the undo journal before its asynchronous workspace operation has fully finished.
            // Undo is available only after the workspace is idle; otherwise the mutation runner rejects it.
            await WaitUntilAsync(() => document.CanUndo && !document.IsWorkspaceBusy);
            Assert.True(document.IsDirty);

            // Undo and redo the edit; moving between tasks keeps the place and zoom.
            await document.UndoCommand.ExecuteAsync(null);
            Assert.False(document.CanUndo);
            await document.RedoCommand.ExecuteAsync(null);
            Assert.True(document.CanUndo);
            document.ShowViewModeCommand.Execute(null);
            Assert.Equal(2, document.SelectedPage!.PageNumber);
            Assert.Equal(zoom, document.Zoom);

            // Save a reviewed copy: the note is in the file and the reader stays where it was.
            await document.SaveAsCommand.ExecuteAsync(null);
            Assert.False(document.IsDirty);
            Assert.Single(PdfDocument.Load(saved).Inspect().GetAnnotationsBySubtype("Text"));
            Assert.Equal(2, document.SelectedPage!.PageNumber);
            Assert.Equal(zoom, document.Zoom);
            return true;
        }, CancellationToken.None);
    }

    private static async Task WaitUntilAsync(Func<bool> condition) {
        for (int attempt = 0; attempt < 200 && !condition(); attempt++) {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(20);
        }
        Assert.True(condition());
    }
}

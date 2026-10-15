using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

/// <summary>Changing the document task keeps the reader's place, zoom, page selection and unsaved work.</summary>
public sealed class StudioModeContinuityTests {
    public static TheoryData<StudioDocumentMode> Modes => new() {
        StudioDocumentMode.Annotate, StudioDocumentMode.Edit, StudioDocumentMode.Pages, StudioDocumentMode.Forms, StudioDocumentMode.Protect
    };

    [Theory]
    [MemberData(nameof(Modes))]
    public async Task Changing_to_a_task_and_back_keeps_place_zoom_selection_and_unsaved_work(StudioDocumentMode mode) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "continuity.pdf");
            PdfDocument.Create(document => {
                for (int page = 1; page <= 8; page++) {
                    int number = page;
                    document.Page(content => content.Size(400, 560).Content(body => body.Item(item => item.Paragraph(text => text.Text("Page " + number)))));
                }
            }).Save(source);
            var window = new MainWindow(services) { Width = 1280, Height = 820 };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(source);
                MainWindowViewModel document = window.ViewModel;
                Layout(window);

                // Unsaved work and a page selection made in the Pages task.
                document.ShowPagesModeCommand.Execute(null);
                document.SetOrganizerSelection([document.OrganizerPages[1], document.OrganizerPages[2]]);
                await document.RotateRightCommand.ExecuteAsync(null);
                Assert.True(document.IsDirty);
                Layout(window);
                document.ShowViewModeCommand.Execute(null);
                Layout(window);
                Assert.Equal(new[] { 2, 3 }, SelectedPages(document));

                // A place in the document at a chosen zoom.
                document.ZoomInCommand.Execute(null);
                double zoom = document.Zoom;
                document.SelectedPage = document.Pages[5];
                Layout(window);
                ScrollViewer reader = ReaderScroll(window);
                double offset = reader.Offset.Y;
                Assert.True(offset > 0);

                document.DocumentMode = mode;
                Layout(window);
                document.DocumentMode = StudioDocumentMode.View;
                Layout(window);

                Assert.Equal(zoom, document.Zoom);
                Assert.Same(document.Pages[5], document.SelectedPage);
                Assert.InRange(ReaderScroll(window).Offset.Y, offset - 2, offset + 2);
                Assert.True(document.IsDirty);
                Assert.True(document.CanUndo);
                Assert.Equal(new[] { 2, 3 }, SelectedPages(document));
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [MemberData(nameof(Modes))]
    public async Task Pending_form_entries_survive_a_change_of_task(StudioDocumentMode mode) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "form-continuity.pdf");
            File.WriteAllBytes(source, PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400)))
                .Forms.Edit(edit => edit.Create(new() { Name = "Reference", Value = "Original", X = 30, Y = 300 })).ToBytes());
            using var document = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await document.OpenDocumentAsync(source);
            document.ShowFormsModeCommand.Execute(null);
            document.SelectedFormField!.TextValue = "Typed but not applied";
            Assert.True(document.HasFormDrafts);

            document.DocumentMode = mode;
            document.ShowFormsModeCommand.Execute(null);

            Assert.True(document.HasFormDrafts);
            Assert.True(document.IsDirty);
            Assert.Equal("Typed but not applied", document.SelectedFormField!.TextValue);
            return true;
        }, CancellationToken.None);
    }

    private static int[] SelectedPages(MainWindowViewModel document) =>
        document.OrganizerPages.Where(page => page.IsSelected).Select(page => page.PageNumber).ToArray();

    private static ScrollViewer ReaderScroll(Window window) =>
        window.GetVisualDescendants().OfType<ListBox>().Single(list => list.Name == "PagesList")
            .GetVisualDescendants().OfType<ScrollViewer>().First();

    private static void Layout(Window window) {
        for (int pass = 0; pass < 3; pass++) {
            Dispatcher.UIThread.RunJobs();
            window.Measure(new Size(1280, 820));
            window.Arrange(new Rect(0, 0, 1280, 820));
            window.UpdateLayout();
        }
        Dispatcher.UIThread.RunJobs();
    }
}

using Avalonia;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Mobile;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed partial class StudioMobileApplicationTests {
    [Theory]
    [InlineData(1366, 1024)]
    [InlineData(390, 844)]
    public async Task FormsAndPageReviewSaveThroughTheMobileHost(int width, int height) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-edit-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                byte[] original = PdfDocument.Create(pdf => {
                    pdf.Page(page => page.Size(300, 400));
                    pdf.Page(page => page.Size(300, 400));
                }).Forms.Edit(edit => edit.Create(new() { Name = "Name", Value = "Original", X = 30, Y = 300 })).ToBytes();
                var source = new TestStorageFile("content://mobile/form.pdf", original, "form.pdf");
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "documents")));
                var view = new MobileWorkspaceView();
                var host = new MobileDocumentHost(services, view, _ => Task.CompletedTask, () => source.CreateProvider());
                using var controller = new MobileDocumentController(services, _ => Task.FromResult<IStorageFile?>(source.Item),
                    _ => Task.CompletedTask, host);
                view.Connect(controller);
                var window = new Window { Content = view, Width = width, Height = height };
                try {
                    window.Show();
                    await controller.Document.OpenCommand.ExecuteAsync(null);
                    var document = controller.Document;
                    await document.Commands["Forms"].ExecuteAsync(); Layout(window);
                    var inspector = Assert.Single(view.GetVisualDescendants().OfType<FormsInspectorView>());
                    var value = inspector.GetVisualDescendants().OfType<TextBox>().Single(box => box.IsEffectivelyVisible && box.Text == "Original");
                    value.Text = "Completed on iPad";
                    Layout(window);
                    Assert.Equal("Completed on iPad", document.SelectedFormField!.TextValue);
                    var apply = inspector.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, document.FillFormFieldCommand));
                    apply.BringIntoView(); Layout(window);
                    Assert.True(apply.Bounds.Height >= 44);
                    Capture(window, $"application-forms-{width}");
                    Click(window, apply);
                    await document.FillFormFieldCommand.ExecutionTask!;
                    Assert.False(document.HasError, document.ErrorMessage);
                    Click(window, view.FindControl<Button>("SaveButton")!);
                    await document.SaveCommand.ExecutionTask!;
                    Assert.False(document.IsDirty, document.ErrorMessage);
                    var field = Assert.Single(PdfDocument.Load(document.DocumentPath!).Inspect().FormFields);
                    Assert.Equal("Completed on iPad", field.Value);
                    Assert.Equal(original, source.Bytes);
                    // Page-tree rewriting has the same PDF capability limits as desktop.
                    // Exercise reordering on a supported document after saving the form.
                    await controller.OpenSampleAsync();
                    document = controller.Document;
                    string[] pageText = PdfReadDocument.Open(document.DocumentPath!).Pages.Select(page => page.ExtractText()).ToArray();
                    await document.Commands["Pages"].ExecuteAsync(); Layout(window);
                    var editor = Assert.Single(view.GetVisualDescendants().OfType<DocumentWorkspaceView>());
                    editor.OrganizerListControl.SelectedItem = document.OrganizerPages[0];
                    Layout(window);
                    using var thumbnailTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                    while (document.OrganizerPages.Any(page => page.IsLoading)) await Task.Delay(10, thumbnailTimeout.Token);
                    Layout(window); Capture(window, $"application-pages-{width}");
                    var actions = editor.FindControl<Border>("OrganizerActionBar")!;
                    foreach (var button in actions.GetVisualDescendants().OfType<Button>().Where(button => button.IsEffectivelyVisible)) {
                        Assert.True(button.Bounds.Width >= 44 && button.Bounds.Height >= 44);
                        var corner = button.TranslatePoint(default, window)!.Value;
                        Assert.InRange(corner.X, 0, window.Bounds.Width - button.Bounds.Width);
                        Assert.InRange(corner.Y, 0, window.Bounds.Height - button.Bounds.Height);
                    }
                    Click(window, view, document.MoveSelectedToCommand);
                    using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                    while (view.FindControl<ContentControl>("DialogContent")!.Content is not PageMoveDialogContent)
                        await Task.Delay(10, timeout.Token);
                    var review = (PageMoveDialogContent)view.FindControl<ContentControl>("DialogContent")!.Content!;
                    review.FindControl<TextBox>("DestinationInput")!.Text = "4";
                    Layout(window); Capture(window, $"application-move-{width}");
                    Click(window, review.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, services.Localizer.Get("Organizer.ApplyMove"))));
                    await document.MoveSelectedToCommand.ExecutionTask!;
                    Assert.False(document.HasError, document.ErrorMessage);
                    Click(window, view.FindControl<Button>("SaveButton")!);
                    await document.SaveCommand.ExecutionTask!;
                    Assert.False(document.IsDirty, document.ErrorMessage);
                    Assert.Equal(new[] { pageText[1], pageText[2], pageText[0] },
                        PdfReadDocument.Open(document.DocumentPath!).Pages.Select(page => page.ExtractText()));
                } finally { window.Close(); services.Storage.Dispose(); }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }
}

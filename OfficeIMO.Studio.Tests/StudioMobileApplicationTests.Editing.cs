using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
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
    [Fact]
    public async Task LandscapePageActionsLeaveRoomForPagesAndDocumentSwitching() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-landscape-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "documents")));
                var view = new MobileWorkspaceView();
                using var controller = new MobileDocumentController(services, _ => Task.FromResult<IStorageFile?>(null), _ => Task.CompletedTask);
                view.Connect(controller);
                var window = new Window { Content = view, Width = 844, Height = 390 };
                try {
                    window.Show();
                    await controller.OpenSampleAsync();
                    await controller.OpenSampleAsync();
                    var document = controller.Document;
                    await document.Commands["Pages"].ExecuteAsync(); Layout(window);
                    var editor = Assert.Single(view.GetVisualDescendants().OfType<DocumentWorkspaceView>());
                    editor.OrganizerListControl.SelectedItem = document.OrganizerPages[1];
                    using var renderTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                    while (document.OrganizerPages.Any(page => page.IsLoading)) await Task.Delay(10, renderTimeout.Token);
                    Layout(window);
                    var actions = editor.FindControl<Border>("OrganizerActionBar")!;
                    var pages = editor.OrganizerListControl;
                    Capture(window, "application-landscape-pages");
                    Assert.True(pages.Bounds.Height >= 100, $"Page viewport is only {pages.Bounds.Height} points high.");
                    double pagesBottom = pages.TranslatePoint(new Point(0, pages.Bounds.Height), window)!.Value.Y;
                    Assert.True(actions.TranslatePoint(default, window)!.Value.Y >= pagesBottom);
                    Click(window, view, document.MoveSelectedUpCommand);
                    await document.MoveSelectedUpCommand.ExecutionTask!;
                    Assert.True(document.IsDirty);
                    Click(window, view.FindControl<Button>("ShortDocumentsButton")!); Layout(window);
                    Assert.True(view.FindControl<ListBox>("DocumentList")!.IsEffectivelyVisible);
                    Assert.Equal(2, view.FindControl<ListBox>("DocumentList")!.ItemCount);
                    var closingTab = controller.Tabs.SelectedTab!;
                    var close = view.FindControl<ListBox>("DocumentList")!.GetVisualDescendants().OfType<Button>()
                        .Single(button => button.Classes.Contains("documentListClose") && ReferenceEquals(button.DataContext, closingTab));
                    Click(window, close); Layout(window);
                    Assert.True(view.FindControl<ScrollViewer>("CloseScroll")!.IsEffectivelyVisible);
                    Click(window, view.FindControl<Button>("CloseCancel")!); Layout(window);
                    await closingTab.CloseCommand.ExecutionTask!;
                    Assert.Equal(2, controller.Tabs.Tabs.Count);
                    Assert.True(document.IsDirty);
                    window.Width = 390; window.Height = 844; Layout(window);
                    Assert.True(view.FindControl<Grid>("TabBar")!.IsEffectivelyVisible);
                    Assert.False(view.FindControl<Button>("ShortDocumentsButton")!.IsEffectivelyVisible);
                    Assert.Same(document, controller.Document);
                    // The opener changes during rotation; dismissal must restore a usable keyboard target.
                    foreach (bool landscape in new[] { true, false }) {
                        Click(window, view.FindControl<Button>(landscape ? "DocumentsButton" : "ShortDocumentsButton")!);
                        window.Width = landscape ? 844 : 390;
                        window.Height = landscape ? 390 : 844;
                        Layout(window);
                        if (landscape) Click(window, view.FindControl<Button>("SheetDone")!);
                        else window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.None, null);
                        Layout(window);
                        var expectedFocus = view.FindControl<Button>(landscape ? "ShortDocumentsButton" : "DocumentsButton")!;
                        Assert.True(ReferenceEquals(expectedFocus, window.FocusManager!.GetFocusedElement()),
                            $"Landscape={landscape}, sheet={view.FindControl<Border>("SheetScrim")!.IsVisible}, targetVisible={expectedFocus.IsEffectivelyVisible}, enabled={expectedFocus.IsEffectivelyEnabled}, focusable={expectedFocus.Focusable}, actual={window.FocusManager.GetFocusedElement()}");
                    }
                    // Compact toolbars retain the shared reading command, including its exit action.
                    var menuButton = editor.FindControl<Button>("DocumentMenuButton")!;
                    Click(window, menuButton); Layout(window);
                    var menu = Assert.IsType<MenuFlyout>(menuButton.Flyout);
                    var focus = Assert.Single(menu.Items.OfType<MenuItem>(),
                        item => ReferenceEquals(item.Command, document.Commands["FocusReading"]));
                    Assert.True(focus.IsEffectivelyVisible && focus.IsEffectivelyEnabled);
                    focus.Command!.Execute(null); menu.Hide(); Layout(window);
                    Assert.True(document.IsFocusReading);
                    Click(window, editor, document.Commands["FocusReading"]); Layout(window);
                    Assert.False(document.IsFocusReading);
                    await document.SelectedPage!.EnsureRenderedAsync();
                    Assert.True(document.SelectedPage.HasScene, "Exiting focus reading must keep the touch reader rendered.");
                    window.Width = 844; window.Height = 390; Layout(window);
                    await document.SelectedPage.EnsureRenderedAsync();
                    Assert.True(document.SelectedPage.HasScene);
                    await document.Commands["Comment"].ExecuteAsync(); Layout(window);
                    document.ShowViewModeCommand.Execute(null); Layout(window);
                    await document.SelectedPage.EnsureRenderedAsync();
                    Assert.True(document.SelectedPage.HasScene, "Returning from editing must keep the touch reader rendered.");
                } finally { window.Close(); services.Storage.Dispose(); }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(1366, 1024)]
    [InlineData(390, 844)]
    [InlineData(844, 390)]
    public async Task FormsAndPageReviewSaveThroughTheMobileHost(int width, int height) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-edit-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                byte[] original = PdfDocument.Create(pdf => {
                    pdf.Page(page => page.Size(300, 400).Content(content => content.Text("Traveller details")));
                    pdf.Page(page => page.Size(300, 400).Content(content => content.Text("Destination details")));
                }).Forms.Edit(edit => edit
                    .Create(new() { Name = "Name", Value = "Original", X = 30, Y = 250 })
                    .Create(new() { Name = "Reference", Value = "STUDIO-123", FieldFlags = 1, X = 30, Y = 190 })
                    .Create(new() { Name = "Destination", PageNumber = 2, Value = "", X = 30, Y = 250 })).ToBytes();
                if (Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT") is { Length: > 0 } evidence) {
                    Directory.CreateDirectory(evidence);
                    File.WriteAllBytes(Path.Combine(evidence, "form-navigation.pdf"), original);
                }
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
                    Assert.False(document.Commands["PreviousFormField"].IsAvailable);
                    Click(window, inspector, document.Commands["NextFormField"]); Layout(window);
                    Assert.Equal("Reference", document.SelectedFormField!.Name);
                    Assert.False(document.SelectedFormField.CanFill);
                    Click(window, inspector, document.Commands["NextFormField"]); Layout(window);
                    Assert.Equal("Destination", document.SelectedFormField!.Name);
                    Assert.Equal(2, document.SelectedPage!.PageNumber);
                    Assert.False(document.Commands["NextFormField"].IsAvailable);
                    var destination = inspector.GetVisualDescendants().OfType<TextBox>()
                        .Single(box => box.IsEffectivelyVisible && Avalonia.Automation.AutomationProperties.GetName(box) == "Destination");
                    destination.Text = "Poland";
                    Layout(window); Capture(window, $"application-form-navigation-{width}");
                    Click(window, inspector, document.Commands["PreviousFormField"]);
                    Click(window, inspector, document.Commands["PreviousFormField"]); Layout(window);
                    Assert.Equal("Name", document.SelectedFormField!.Name);
                    Assert.Equal(1, document.SelectedPage!.PageNumber);
                    Assert.Equal("Completed on iPad", document.SelectedFormField.TextValue);
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
                    var fields = PdfDocument.Load(document.DocumentPath!).Inspect().FormFields;
                    Assert.Equal("Completed on iPad", fields.Single(field => field.Name == "Name").Value);
                    Assert.Equal("Poland", fields.Single(field => field.Name == "Destination").Value);
                    Assert.Equal("STUDIO-123", fields.Single(field => field.Name == "Reference").Value);
                    Assert.Equal(original, source.Bytes);
                    // Page-tree rewriting has the same PDF capability limits as desktop.
                    // Exercise reordering on a supported document after saving the form.
                    await controller.OpenSampleAsync();
                    document = controller.Document;
                    Assert.False(document.Commands["PreviousFormField"].IsAvailable);
                    Assert.False(document.Commands["NextFormField"].IsAvailable);
                    string[] pageText = PdfReadDocument.Open(document.DocumentPath!).Pages.Select(page => page.ExtractText()).ToArray();
                    await document.Commands["Pages"].ExecuteAsync(); Layout(window);
                    var editor = Assert.Single(view.GetVisualDescendants().OfType<DocumentWorkspaceView>());
                    editor.OrganizerListControl.SelectedItem = document.OrganizerPages[0];
                    Layout(window);
                    using var thumbnailTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                    while (document.OrganizerPages.Any(page => page.IsLoading)) await Task.Delay(10, thumbnailTimeout.Token);
                    Layout(window); Capture(window, $"application-pages-{width}");
                    Assert.All(editor.OrganizerListControl.GetRealizedContainers(), page => Assert.InRange(page.Bounds.Width, 44, 250));
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

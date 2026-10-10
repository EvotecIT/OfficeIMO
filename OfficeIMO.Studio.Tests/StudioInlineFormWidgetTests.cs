using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioInlineFormWidgetTests {
    [Fact]
    public async Task RadioAndListControlsKeepDraftsThroughNavigationSaveAndUndo() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "choices.pdf"), output = Path.Combine(services.Paths.Root, "choices-filled.pdf");
            PdfDocument.Create(compose => compose.Page(page => page.Size(420, 520).Margin(30).Content(content => content
                .Paragraph(p => p.Text("Project preferences"))
                .Paragraph(p => p.Text("Delivery method"))
                .RadioButtonGroup("Delivery", ["Email", "Post", "Collect"], value: "Email", size: 24, gap: 12)
                .Paragraph(p => p.Text("Region"))
                .ChoiceField("Region", ["Europe", "Americas", "Asia Pacific"], "Europe", height: 112, isComboBox: false)
                .Paragraph(p => p.Text("Topics (select more than one)"))
                .MultiSelectChoiceField("Topics", ["Design", "Engineering", "Research"], ["Design"], height: 112))))
                .Save(source);
            string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
            if (!string.IsNullOrEmpty(evidence)) {
                Directory.CreateDirectory(evidence);
                File.Copy(source, Path.Combine(evidence, "inline-forms-source.pdf"), true);
            }
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), pickSavePdf: _ => Task.FromResult<string?>(output), services: services);
            await model.OpenDocumentAsync(source);
            model.SetViewportSize(600, 800);
            model.ShowFormsModeCommand.Execute(null);
            var page = model.SelectedPage!;
            var window = new Window { Width = 1280, Height = 900, RequestedThemeVariant = ThemeVariant.Dark,
                Content = new DocumentWorkspaceView { DataContext = model } };
            try {
                window.Show();
                model.SelectedFormField = model.FormFields.Single(field => field.Name == "Delivery");
                await page.EnsureRenderedAsync();
                await Layout(window);
                var radios = window.GetVisualDescendants().OfType<RadioButton>().Where(control => control.IsEffectivelyVisible).ToArray();
                Assert.Equal(3, radios.Length);
                Assert.True(radios[0].IsChecked);
                await Click(window, radios[1]);
                Assert.Equal("Post", model.SelectedFormField.SelectedChoice?.ExportValue);
                Assert.False(radios[0].IsChecked);
                Assert.True(radios[1].IsChecked);
                window.KeyPress(Key.Right, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Equal("Collect", model.SelectedFormField.SelectedChoice?.ExportValue);
                Assert.True(radios[2].IsFocused);
                window.KeyPress(Key.Tab, RawInputModifiers.None, PhysicalKey.None, null);
                await Layout(window);
                Assert.Equal("Region", model.SelectedFormField.Name);
                Assert.Equal("Collect", model.FormFields.Single(field => field.Name == "Delivery").SelectedChoice?.ExportValue);
                var list = VisibleList(window);
                Assert.Equal("Europe", Assert.IsType<PdfFormChoiceViewModel>(Assert.Single(list.SelectedItems!)).ExportValue);
                Assert.True(list.IsKeyboardFocusWithin);
                window.KeyPress(Key.Down, RawInputModifiers.None, PhysicalKey.None, null);
                await Layout(window);
                Assert.Equal("Americas", model.SelectedFormField.SelectedChoice?.ExportValue);
                var inspectorChoice = window.GetVisualDescendants().OfType<FormsInspectorView>().Single(view => view.IsEffectivelyVisible)
                    .GetVisualDescendants().OfType<ComboBox>()
                    .Single(control => control.IsEffectivelyVisible && ReferenceEquals(control.ItemsSource, model.SelectedFormField.Choices));
                inspectorChoice.SelectedItem = model.SelectedFormField.Choices[0];
                Assert.Equal("Europe", model.SelectedFormField.SelectedChoice?.ExportValue);
                inspectorChoice.SelectedItem = model.SelectedFormField.Choices[1];
                Assert.Equal("Americas", Assert.IsType<PdfFormChoiceViewModel>(list.SelectedItem).ExportValue);
                window.KeyPress(Key.Tab, RawInputModifiers.None, PhysicalKey.None, null);
                await Layout(window);
                Assert.Equal("Topics", model.SelectedFormField.Name);
                list = VisibleList(window);
                Assert.Equal("Design", Assert.IsType<PdfFormChoiceViewModel>(Assert.Single(list.SelectedItems!)).ExportValue);
                await Click(window, list.GetVisualDescendants().OfType<ListBoxItem>().ElementAt(2));
                Assert.Equal(new[] { "Design", "Research" }, model.SelectedFormField.CreateValue().Values);
                // Recreating the list after field navigation must not replace the draft with its default selection.
                window.KeyPress(Key.Tab, RawInputModifiers.Shift, PhysicalKey.None, null);
                await Layout(window);
                Assert.Equal("Region", model.SelectedFormField.Name);
                Assert.True(VisibleList(window).IsKeyboardFocusWithin);
                window.KeyPress(Key.Tab, RawInputModifiers.None, PhysicalKey.None, null);
                await Layout(window);
                Assert.Equal("Topics", model.SelectedFormField.Name);
                Assert.Equal(2, VisibleList(window).SelectedItems!.Count);
                Assert.True(radios[2].IsChecked); // Moving to a different field keeps the unsaved choice visible.
                if (!string.IsNullOrEmpty(evidence)) {
                    using var image = window.CaptureRenderedFrame();
                    image!.Save(Path.Combine(evidence, "inline-forms-headless-list.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
                await model.SaveAsCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage);
                Assert.False(model.HasFormDrafts);
                var saved = PdfDocument.Load(output).Inspect().FormFields;
                var radio = saved.Single(field => field.Name == "Delivery");
                Assert.Equal("Collect", radio.Value);
                Assert.Single(radio.Widgets, widget => widget.AppearanceState == "Collect");
                Assert.Equal("Americas", saved.Single(field => field.Name == "Region").Value);
                Assert.Equal(new[] { "Design", "Research" }, saved.Single(field => field.Name == "Topics").Values);
                await model.UndoCommand.ExecuteAsync(null);
                Assert.Equal("Email", model.FormFields.Single(field => field.Name == "Delivery").SelectedChoice?.ExportValue);
                Assert.Equal(new[] { "Design" }, model.FormFields.Single(field => field.Name == "Topics").CreateValue().Values);
                if (!string.IsNullOrEmpty(evidence)) File.Copy(output, Path.Combine(evidence, "inline-forms-headless-saved.pdf"), true);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task SingleListClearingSelectionSavesAnEmptyValue() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "single-list.pdf"), output = Path.Combine(services.Paths.Root, "cleared.pdf");
            PdfDocument.Create(compose => compose.Page(page => page.Size(420, 520).Margin(30).Content(content => content
                .ChoiceField("Region", ["Europe", "Americas", "Asia Pacific"], "Europe", height: 112, isComboBox: false)))).Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), pickSavePdf: _ => Task.FromResult<string?>(output), services: services);
            await model.OpenDocumentAsync(source);
            model.SetViewportSize(600, 800);
            model.ShowFormsModeCommand.Execute(null);
            var page = model.SelectedPage!;
            var window = new Window { Width = 760, Height = 900, Content = new PdfPageView { DataContext = page } };
            try {
                window.Show();
                await page.EnsureRenderedAsync();
                await Layout(window);
                var list = VisibleList(window);
                var selected = list.GetVisualDescendants().OfType<ListBoxItem>().Single(item => item.IsSelected);
                await Click(window, selected, RawInputModifiers.Control);
                Assert.Empty(list.SelectedItems!);
                Assert.Null(model.SelectedFormField!.SelectedChoice);
                Assert.True(model.HasFormDrafts);
                await model.SaveAsCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage);
                Assert.True(string.IsNullOrEmpty(PdfDocument.Load(output).Inspect().FormFields.Single().Value));
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task LongListSelectAllSavesOffscreenChoicesAndRestoresThemAfterRecreation() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "long-list.pdf"), output = Path.Combine(services.Paths.Root, "all-selected.pdf");
            string[] choices = Enumerable.Range(1, 40).Select(index => $"Choice {index:00}").ToArray();
            PdfDocument.Create(compose => compose.Page(page => page.Size(420, 520).Margin(30).Content(content => content
                .MultiSelectChoiceField("Topics", choices, [choices[^1]], height: 96)))).Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), pickSavePdf: _ => Task.FromResult<string?>(output), services: services);
            await model.OpenDocumentAsync(source);
            model.SetViewportSize(600, 800);
            model.ShowFormsModeCommand.Execute(null);
            var page = model.SelectedPage!;
            var window = new Window { Width = 760, Height = 900, Content = new PdfPageView { DataContext = page } };
            try {
                window.Show();
                await page.EnsureRenderedAsync();
                await Layout(window);
                var list = VisibleList(window);
                Assert.True(list.GetVisualDescendants().OfType<ListBoxItem>().Count() < choices.Length);
                Assert.Equal(choices[^1], Assert.IsType<PdfFormChoiceViewModel>(Assert.Single(list.SelectedItems!)).ExportValue);
                await Click(window, list.GetVisualDescendants().OfType<ListBoxItem>().First());
                window.KeyPress(Key.A, RawInputModifiers.Control, PhysicalKey.None, null);
                await Layout(window);
                Assert.Equal(choices.Length, list.SelectedItems!.Count);
                Assert.Equal(choices, model.SelectedFormField!.CreateValue().Values);
                // A recreated control must seed its complete selection, including virtualized rows.
                window.Content = new PdfPageView { DataContext = page };
                await Layout(window);
                Assert.Equal(choices.Length, VisibleList(window).SelectedItems!.Count);
                await model.SaveAsCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage);
                Assert.Equal(choices, PdfDocument.Load(output).Inspect().FormFields.Single().Values);
                await model.UndoCommand.ExecuteAsync(null);
                Assert.Equal(new[] { choices[^1] }, model.SelectedFormField!.CreateValue().Values);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static ListBox VisibleList(Window window) => window.GetVisualDescendants().OfType<ListBox>()
        .Single(control => control.Name == "InlineFormList" && control.IsEffectivelyVisible &&
            control.DataContext is PdfInlineFormWidgetViewModel widget && ReferenceEquals(widget.Field, widget.Page.InlineFormField));

    private static async Task Layout(Window window) {
        await Dispatcher.UIThread.InvokeAsync(window.UpdateLayout, DispatcherPriority.Background);
        AvaloniaHeadlessPlatform.ForceRenderTimerTick();
    }

    private static async Task Click(Window window, Control control, RawInputModifiers modifiers = RawInputModifiers.None) {
        control.BringIntoView();
        await Layout(window);
        Point point = control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;
        window.MouseDown(point, MouseButton.Left, modifiers);
        window.MouseUp(point, MouseButton.Left, modifiers);
        await Layout(window);
    }
}

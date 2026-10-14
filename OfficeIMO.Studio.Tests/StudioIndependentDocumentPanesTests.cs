using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Styling;
using Avalonia.VisualTree;
using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioIndependentDocumentPanesTests {
    [Theory]
    [InlineData(1600, false, "en")]
    [InlineData(960, true, "en")]
    [InlineData(390, false, "en")]
    [InlineData(390, false, "pl")]
    [InlineData(390, true, "de")]
    [InlineData(1600, false, "fr")]
    public async Task SameDocumentHasIndependentNavigationAndSharedEditingUndoAndSave(int width, bool dark, string culture) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            Application.Current!.RequestedThemeVariant = dark ? ThemeVariant.Dark : ThemeVariant.Light;
            IStudioLocalizer originalLocalizer = StudioLocalization.Current;
            var localizer = new StudioLocalizer(CultureInfo.GetCultureInfo(culture));
            StudioLocalization.Configure(localizer);
            var services = TestAppBuilder.CreateTestServices();
            services.Preferences.Update(preferences => preferences with { UiCulture = culture });
            services = StudioApplicationServices.Create(services.Paths);
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "Shared three-page document.pdf");
            string saved = Path.Combine(services.Paths.Root, "shared-saved.pdf");
            Create(source, "Shared", 3);
            byte[] original = File.ReadAllBytes(source);
            var view = new DocumentWorkspaceView();
            using var host = new StudioDocumentTabHost(open => new MainWindowViewModel(
                _ => Task.FromResult<string?>(null), pickSavePdf: _ => Task.FromResult<string?>(saved),
                openDocumentInTab: open, services: services), document => view.DataContext = document);
            var window = new Window { Width = width, Height = 900, Content = view };
            view.Panes = host.Panes;
            try {
                window.Show();
                await host.OpenDocumentAsync(source);
                var document = host.ActiveDocument;
                document.SelectedReaderLayoutChoice = document.ReaderLayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.SinglePage);
                host.Panes.OpenSecondPane(host.SelectedTab);
                var left = host.Panes.Left!; var right = host.Panes.Right!;
                left.SelectedLayout = left.LayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.SinglePage);
                left.Zoom = .75;
                right.SelectedLayout = right.LayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.Continuous);
                right.ActivatePage(3); right.Zoom = .5;
                Assert.Same(left.Document, right.Document);
                Assert.Single(host.Tabs);
                Assert.Equal(1, left.SelectedPage!.PageNumber);
                Assert.Equal(3, right.SelectedPage!.PageNumber);
                Assert.Equal(.75, left.Zoom);
                Assert.Equal(.5, right.Zoom);
                Assert.NotSame(left.Pages[0], right.Pages[0]);
                document.PreviousPageCommand.Execute(null);
                Assert.Equal(2, right.SelectedPage!.PageNumber); Assert.Equal(1, left.SelectedPage!.PageNumber);
                document.NextPageCommand.Execute(null);
                document.SelectedReaderLayoutChoice = document.ReaderLayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.TwoPage);
                Assert.Equal(ReaderLayoutMode.TwoPage, right.SelectedLayout.Mode);
                Assert.Equal(ReaderLayoutMode.SinglePage, left.SelectedLayout.Mode);
                document.SelectedReaderLayoutChoice = document.ReaderLayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.Continuous);
                right.Zoom = .5;
                await Render(window, left, right);
                Assert.Equal(.75, left.Zoom); Assert.Equal(.5, right.Zoom);
                Assert.NotNull(right.SelectedPage!.Scene);
                if (width == 1600) Assert.NotNull(left.SelectedPage!.Scene);
                Capture(window, $"panes-same-{width}-{(dark ? "dark" : "light")}-{culture}");
                document.ShowAnnotateModeCommand.Execute(null);
                document.SelectedEditorToolChoice = document.EditorTools.Single(choice => choice.Tool == PdfEditorTool.Note);
                right.Pages[2].CompleteEditorGesture(new(3, 70, 80, 95, 105, []));
                await WaitForMutation(document);
                Assert.True(document.IsDirty, document.ErrorMessage);
                Assert.Equal(3, left.Pages.Count); Assert.Equal(3, right.Pages.Count);
                Assert.Equal(1, left.SelectedPage!.PageNumber); Assert.Equal(3, right.SelectedPage!.PageNumber);
                await document.SaveAsCommand.ExecuteAsync(null);
                var changed = PdfDocument.Load(saved).Inspect();
                Assert.Single(changed.Annotations, annotation => annotation.Subtype == "Text");
                Assert.Equal(3, changed.Annotations.Single(annotation => annotation.Subtype == "Text").PageNumber);
                Assert.Equal(original, File.ReadAllBytes(source));
                var note = changed.Annotations.Single(annotation => annotation.Subtype == "Text");
                right.Pages[2].SelectObject(new(PdfEditorSelectionKind.Annotation, 3, new(70, 80, 95, 105),
                    ObjectNumber: note.ObjectNumber, Subtype: note.Subtype));
                Assert.Single(document.SelectedAnnotations);
                Assert.Equal(note.ObjectNumber, left.Pages[2].SelectedObject?.ObjectNumber);
                Assert.Equal(note.ObjectNumber, right.Pages[2].SelectedObject?.ObjectNumber);
                left.Activate();
                Assert.Equal(1, document.SelectedPage!.PageNumber);
                Assert.Equal(.75, document.Zoom);
                Assert.Equal(note.ObjectNumber, Assert.Single(document.SelectedAnnotations).ObjectNumber);
                await document.UndoCommand.ExecuteAsync(null);
                await document.SaveAsCommand.ExecuteAsync(null);
                Assert.Empty(PdfDocument.Load(saved).Inspect().Annotations);
                Assert.Equal(1, left.SelectedPage!.PageNumber); Assert.Equal(3, right.SelectedPage!.PageNumber);
                Assert.Equal(ReaderLayoutMode.SinglePage, left.SelectedLayout.Mode);
                Assert.Equal(ReaderLayoutMode.Continuous, right.SelectedLayout.Mode);
                await Render(window, left, right);
                Assert.Equal(.75, left.Zoom); Assert.Equal(.5, right.Zoom);
                Capture(window, $"panes-undo-{width}-{culture}");
                host.Panes.ClosePane(left);
                Assert.False(host.Panes.IsSplit);
                Assert.Single(host.Tabs);
                Assert.True(document.HasDocument);
                Assert.Equal(3, document.SelectedPage!.PageNumber);
            } finally { window.Close(); StudioLocalization.Configure(originalLocalizer); }
            return true;
        }, default);
    }

    [Fact]
    public async Task TwoDocumentsRouteFocusSelectionUndoSaveAndUnsavedCloseToTheirOwnSession() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string first = Path.Combine(services.Paths.Root, "First.pdf"), second = Path.Combine(services.Paths.Root, "Second.pdf");
            string firstSaved = Path.Combine(services.Paths.Root, "First-saved.pdf"), secondSaved = Path.Combine(services.Paths.Root, "Second-saved.pdf");
            Create(first, "First", 2); Create(second, "Second", 3);
            int prompts = 0;
            UnsavedChangesDecision decision = UnsavedChangesDecision.Cancel;
            var view = new DocumentWorkspaceView();
            MainWindowViewModel? active = null;
            using var host = new StudioDocumentTabHost(open => {
                MainWindowViewModel? model = null;
                model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                    pickSavePdf: _ => Task.FromResult<string?>(model!.DocumentName.StartsWith("First", StringComparison.Ordinal) ? firstSaved : secondSaved),
                    confirmUnsavedChanges: () => { prompts++; return Task.FromResult(decision); }, openDocumentInTab: open, services: services);
                return model;
            }, document => { active = document; view.DataContext = document; });
            view.Panes = host.Panes;
            var window = new Window { Width = 1600, Height = 900, Content = view };
            try {
                window.Show();
                await host.OpenDocumentAsync(first); var firstTab = host.SelectedTab!;
                host.Panes.OpenSecondPane(firstTab);
                await host.OpenDocumentAsync(second); var secondTab = host.SelectedTab!;
                var left = host.Panes.Left!; var right = host.Panes.Right!;
                Assert.Same(firstTab.Document, left.Document); Assert.Same(secondTab.Document, right.Document);
                left.ActivatePage(2); left.Zoom = .6;
                right.ActivatePage(3); right.Zoom = .9;
                await Render(window, left, right);
                Assert.Equal(.6, left.Zoom); Assert.Equal(.9, right.Zoom);
                Assert.NotNull(left.SelectedPage!.Scene); Assert.NotNull(right.SelectedPage!.Scene);
                Capture(window, "panes-two-documents-wide");
                string invalid = Path.Combine(services.Paths.Root, "missing.pdf");
                await host.OpenDocumentAsync(invalid);
                Assert.Equal(2, host.Tabs.Count);
                Assert.True(host.Panes.IsSplit);
                Assert.Same(left, host.Panes.Left); Assert.Same(right, host.Panes.Right);
                Assert.Same(secondTab, host.SelectedTab);
                Assert.Equal(2, left.SelectedPage!.PageNumber); Assert.Equal(3, right.SelectedPage!.PageNumber);
                Assert.NotNull(right.Document.ErrorMessage);
                right.Document.ErrorMessage = null;
                foreach (var pane in new[] { left, right }) {
                    pane.Document.ShowAnnotateModeCommand.Execute(null);
                    pane.Document.SelectedEditorToolChoice = pane.Document.EditorTools.Single(choice => choice.Tool == PdfEditorTool.Note);
                    pane.Pages[pane.SelectedPage!.PageNumber - 1].CompleteEditorGesture(new(pane.SelectedPage.PageNumber, 80, 80, 105, 105, []));
                    await WaitForMutation(pane.Document);
                    Assert.Same(pane.Document, active);
                    Assert.Same(pane.Tab, host.SelectedTab);
                }
                Assert.True(firstTab.Document.IsDirty); Assert.True(secondTab.Document.IsDirty);
                left.Activate();
                await active!.UndoCommand.ExecuteAsync(null);
                Assert.False(firstTab.Document.IsDirty); Assert.True(secondTab.Document.IsDirty);
                await active.SaveAsCommand.ExecuteAsync(null);
                Assert.Empty(PdfDocument.Load(firstSaved).Inspect().Annotations);
                right.Activate();
                await active!.SaveAsCommand.ExecuteAsync(null);
                var savedNote = Assert.Single(PdfDocument.Load(secondSaved).Inspect().Annotations, annotation => annotation.Subtype == "Text");
                right.Pages[2].SelectObject(new(PdfEditorSelectionKind.Annotation, 3, new(80, 80, 105, 105),
                    ObjectNumber: savedNote.ObjectNumber, Subtype: savedNote.Subtype));
                left.Activate();
                Assert.Empty(active!.SelectedAnnotations);
                Assert.Equal(savedNote.ObjectNumber, Assert.Single(right.Document.SelectedAnnotations).ObjectNumber);
                right.Activate();
                Assert.Equal(savedNote.ObjectNumber, Assert.Single(active!.SelectedAnnotations).ObjectNumber);
                right.Pages[2].CompleteEditorGesture(new(3, 120, 80, 145, 105, []));
                await WaitForMutation(right.Document);
                await host.CloseTabAsync(secondTab);
                Assert.Equal(1, prompts);
                Assert.Contains(secondTab, host.Tabs); Assert.True(host.Panes.IsSplit);
                Assert.Equal(2, left.SelectedPage!.PageNumber); Assert.Equal(3, right.SelectedPage!.PageNumber);
                decision = UnsavedChangesDecision.Discard;
                await host.CloseTabAsync(secondTab);
                Assert.Equal(2, prompts);
                Assert.DoesNotContain(secondTab, host.Tabs); Assert.False(host.Panes.IsSplit);
                Assert.Same(firstTab.Document, host.ActiveDocument);
                Assert.True(firstTab.Document.HasDocument);
            } finally { window.Close(); }
            return true;
        }, default);
    }

    [Theory]
    [InlineData(1600)]
    [InlineData(390)]
    public async Task ShellF6AndComparisonKeepIndependentPaneModeDistinct(int width) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "Current.pdf"), other = Path.Combine(services.Paths.Root, "Other.pdf");
            Create(source, "Current", 3); Create(other, "Other", 3);
            var window = new MainWindow(services) { Width = width, Height = 900 };
            try {
                window.Show(); await window.TabHost.OpenDocumentAsync(source);
                window.TabHost.Panes.OpenSecondPane(window.TabHost.SelectedTab);
                window.UpdateLayout();
                var panes = window.TabHost.Panes;
                panes.Left!.SelectedPage = panes.Left.Pages[0]; panes.Right!.ActivatePage(3);
                window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!.FocusActiveIndependentPane();
                window.KeyPress(Key.F6, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Same(panes.Left, panes.ActivePane);
                Assert.Equal(1, window.ViewModel.SelectedPage!.PageNumber);
                window.KeyPress(Key.F6, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Same(panes.Right, panes.ActivePane);
                Assert.Equal(3, window.ViewModel.SelectedPage!.PageNumber);
                var splitView = window.GetVisualDescendants().OfType<IndependentDocumentPanesView>().Single();
                var content = window.Content;
                window.Content = null;
                panes.SwitchPane();
                window.Content = content;
                window.UpdateLayout();
                splitView.FocusActivePane();
                window.KeyPress(Key.F6, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Same(panes.Right, panes.ActivePane);
                Assert.Equal(width < 780, splitView.IsCompact);
                Assert.Equal(width >= 780, splitView.FindControl<IndependentDocumentPaneView>("LeftPane")!.IsVisible);
                Assert.True(splitView.FindControl<IndependentDocumentPaneView>("RightPane")!.IsVisible);
                window.ViewModel.ShowHomeCommand.Execute(null);
                window.UpdateLayout();
                Assert.False(splitView.IsEffectivelyVisible);
                window.ViewModel.ShowPdfWorkspaceCommand.Execute(null);
                window.UpdateLayout();
                Assert.True(splitView.IsEffectivelyVisible);
                Assert.Same(panes.Right, panes.ActivePane);
                await Render(window, panes.Left, panes.Right);
                if (width < 780) Assert.InRange(panes.Right.Zoom, .25, 1.3);
                Capture(window, $"panes-reattached-{width}");
                await window.ViewModel.OpenComparisonDocumentAsync(other);
                window.UpdateLayout();
                Assert.True(window.ViewModel.IsComparisonOpen);
                Assert.False(splitView.IsEffectivelyVisible);
                window.ViewModel.NextPageCommand.Execute(null);
                Assert.Equal(window.ViewModel.SelectedPage!.PageNumber, window.ViewModel.ComparisonSelectedPage!.PageNumber);
                window.ViewModel.CloseComparisonCommand.Execute(null);
                window.UpdateLayout();
                Assert.True(splitView.IsEffectivelyVisible);
                Assert.Equal(1, panes.Left.SelectedPage!.PageNumber);
                Assert.Equal(3, panes.Right.SelectedPage!.PageNumber);
                await window.TabHost.OpenDocumentAsync(other);
                Assert.Equal(other, window.ViewModel.DocumentPath);
                window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!.FocusActiveIndependentPane();
                window.KeyPress(Key.F6, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Equal(source, window.ViewModel.DocumentPath);
                Assert.Same(panes.Left, panes.ActivePane);
                window.KeyPress(Key.F6, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Equal(other, window.ViewModel.DocumentPath);
                Assert.Same(panes.Right, panes.ActivePane);
                panes.ClosePane(panes.Right!);
                Assert.False(panes.IsSplit);
                Assert.Equal(source, window.ViewModel.DocumentPath);
                Assert.Equal(2, window.TabHost.Tabs.Count);
            } finally { window.Close(); }
            return true;
        }, default);
    }

    [Fact]
    public async Task SplittingComparisonRestoresReaderLayoutBeforeCreatingIndependentPanes() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "Split source.pdf");
            string comparison = Path.Combine(services.Paths.Root, "Split comparison.pdf");
            Create(source, "Split source", 3); Create(comparison, "Split comparison", 3);
            var window = new MainWindow(services) { Width = 1600, Height = 900 };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(source);
                window.ViewModel.SelectedReaderLayoutChoice = window.ViewModel.ReaderLayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.TwoPage);
                await window.ViewModel.OpenComparisonDocumentAsync(comparison);
                Assert.True(window.ViewModel.IsComparisonOpen, window.ViewModel.ErrorMessage);
                window.TabHost.Panes.OpenSecondPane(window.TabHost.SelectedTab);
                var panes = window.TabHost.Panes;
                await Render(window, panes.Left!, panes.Right!);
                Capture(window, "panes-split-comparison-restored-layout");
                Assert.False(window.ViewModel.IsComparisonOpen);
                Assert.Empty(window.ViewModel.ComparisonPages);
                Assert.True(panes.IsSplit);
                Assert.Equal(ReaderLayoutMode.TwoPage, panes.Left!.SelectedLayout.Mode);
                Assert.Equal(ReaderLayoutMode.TwoPage, panes.Right!.SelectedLayout.Mode);
                Assert.Equal(ReaderLayoutMode.TwoPage, window.ViewModel.SelectedReaderLayoutChoice.Mode);
                Assert.True(window.GetVisualDescendants().OfType<IndependentDocumentPanesView>().Single().IsEffectivelyVisible);
            } finally { window.Close(); }
            return true;
        }, default);
    }

    [Theory]
    [InlineData(1600, false)]
    [InlineData(1600, true)]
    [InlineData(390, false)]
    [InlineData(390, true)]
    public async Task AssigningComparisonTabRestoresIndependentPanesAndItsReaderLayout(int width, bool selectGlobalTab) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string first = Path.Combine(services.Paths.Root, "First pane.pdf");
            string replacement = Path.Combine(services.Paths.Root, "Replacement pane.pdf");
            string comparison = Path.Combine(services.Paths.Root, "Comparison partner.pdf");
            Create(first, "First pane", 3); Create(replacement, "Replacement pane", 3); Create(comparison, "Comparison partner", 3);
            byte[] originalFirst = File.ReadAllBytes(first), originalReplacement = File.ReadAllBytes(replacement);
            var window = new MainWindow(services) { Width = width, Height = 900 };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(first);
                var firstTab = window.TabHost.SelectedTab!;
                await window.TabHost.OpenDocumentAsync(replacement);
                var replacementTab = window.TabHost.SelectedTab!;
                replacementTab.Document.SelectedReaderLayoutChoice = replacementTab.Document.ReaderLayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.TwoPage);
                replacementTab.Document.SelectedPage = replacementTab.Document.Pages[1];
                await replacementTab.Document.OpenComparisonDocumentAsync(comparison);
                Assert.True(replacementTab.Document.IsComparisonOpen, replacementTab.Document.ErrorMessage);
                Assert.Equal(ReaderLayoutMode.SinglePage, replacementTab.Document.SelectedReaderLayoutChoice.Mode);
                window.TabHost.SelectedTab = firstTab;
                window.TabHost.Panes.OpenSecondPane(firstTab);
                var panes = window.TabHost.Panes;
                var left = panes.Left!;
                var previousRight = panes.Right!;
                left.ActivatePage(3); left.Zoom = .6;
                previousRight.Activate();

                if (selectGlobalTab) window.TabHost.SelectedTab = replacementTab;
                else previousRight.SelectedTab = replacementTab;

                var right = panes.Right!;
                await Render(window, left, right);
                var splitView = window.GetVisualDescendants().OfType<IndependentDocumentPanesView>().Single();
                Capture(window, $"panes-assigned-comparison-{width}-{(selectGlobalTab ? "tab" : "pane")}");
                Assert.False(replacementTab.Document.IsComparisonOpen);
                Assert.Empty(replacementTab.Document.ComparisonPages);
                Assert.True(panes.IsSplit);
                Assert.Same(left, panes.Left);
                Assert.NotSame(previousRight, right);
                Assert.Same(replacementTab.Document, right.Document);
                Assert.Same(right, panes.ActivePane);
                Assert.Equal(ReaderLayoutMode.TwoPage, right.SelectedLayout.Mode);
                Assert.Equal(2, right.SelectedPage!.PageNumber);
                Assert.Equal(3, left.SelectedPage!.PageNumber);
                Assert.Equal(.6, left.Zoom);
                Assert.True(splitView.IsEffectivelyVisible);
                Assert.Equal(width >= 780, splitView.FindControl<IndependentDocumentPaneView>("LeftPane")!.IsVisible);
                Assert.True(splitView.FindControl<IndependentDocumentPaneView>("RightPane")!.IsVisible);
                var workspace = window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!;
                workspace.FocusActiveIndependentPane();
                window.KeyPress(Key.F6, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Same(left, panes.ActivePane);
                Assert.Same(firstTab.Document, window.ViewModel);
                await Render(window, left, right);
                Capture(window, $"panes-assigned-comparison-f6-left-{width}-{(selectGlobalTab ? "tab" : "pane")}");
                Assert.Equal(3, left.SelectedPage!.PageNumber);
                Assert.Equal(3, window.ViewModel.SelectedPage!.PageNumber);
                window.KeyPress(Key.F6, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Same(right, panes.ActivePane);
                Assert.Same(replacementTab.Document, window.ViewModel);
                await Render(window, left, right);
                Capture(window, $"panes-assigned-comparison-f6-right-{width}-{(selectGlobalTab ? "tab" : "pane")}");
                Assert.Equal(2, right.SelectedPage!.PageNumber);
                Assert.Equal(2, window.ViewModel.SelectedPage!.PageNumber);
                Assert.Equal(originalFirst, File.ReadAllBytes(first));
                Assert.Equal(originalReplacement, File.ReadAllBytes(replacement));
            } finally { window.Close(); }
            return true;
        }, default);
    }

    private static async Task Render(Window window, params StudioDocumentPaneViewModel[] panes) {
        window.Measure(new Size(window.Width, window.Height)); window.Arrange(new Rect(0, 0, window.Width, window.Height)); window.UpdateLayout();
        for (int pass = 0; pass < 2; pass++) {
            foreach (var pane in panes) foreach (var page in pane.Pages) await page.EnsureRenderedAsync();
            window.UpdateLayout();
        }
        for (int attempt = 0; attempt < 200 && panes.SelectMany(pane => pane.Pages).Any(page => page.IsRendering); attempt++)
            await Task.Delay(10);
        Assert.DoesNotContain(panes.SelectMany(pane => pane.Pages), page => page.IsRendering);
    }
    private static async Task WaitForMutation(MainWindowViewModel document) {
        for (int attempt = 0; attempt < 200 && document.IsWorkspaceBusy; attempt++) await Task.Delay(10);
        Assert.False(document.IsWorkspaceBusy); Assert.Null(document.ErrorMessage);
    }
    private static void Create(string path, string title, int pages) => PdfDocument.Create(document => {
        for (int page = 1; page <= pages; page++) {
            int number = page;
            document.Page(p => p.Size(300, 400).Content(content => content.Item(item => item.Paragraph(text => text.Text($"{title} · page {number}")))));
        }
    }).Save(path);
    private static void Capture(Window window, string name) {
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output); frame.Save(Path.Combine(output, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}

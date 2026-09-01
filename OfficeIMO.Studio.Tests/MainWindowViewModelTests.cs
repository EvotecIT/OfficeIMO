using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Tests;

public sealed class MainWindowViewModelTests {
    [Fact]
    public async Task RedactionPreviewPersistsOnPageUntilCancelled() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-redaction-preview-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "editable.pdf");
        PdfDocument.Create(compose => compose.Page(page => page
            .Size(600D, 800D)
            .Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Sensitive review text"))))))
            .Save(path);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);
            viewModel.SelectedEditorToolChoice = viewModel.EditorTools.Single(choice => choice.Tool == PdfEditorTool.Redact);
            var gesture = new PdfEditorGesture(
                1,
                36D,
                48D,
                240D,
                92D,
                Array.Empty<PdfEditorVisualPoint>());

            viewModel.Pages[0].CompleteEditorGesture(gesture);
            await WaitUntilAsync(() => viewModel.HasPendingRedaction);

            Assert.Equal(new Avalonia.Rect(36D, 48D, 204D, 44D), viewModel.Pages[0].PendingRedactionArea);

            viewModel.CancelPendingRedactionCommand.Execute(null);

            Assert.False(viewModel.HasPendingRedaction);
            Assert.Null(viewModel.Pages[0].PendingRedactionArea);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task SuccessfulMutationInvalidatesReviewedRedactionAndAnnotationSelection() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-review-state-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "editable.pdf");
        byte[] source = PdfDocument.Create(compose => compose.Page(page => page
            .Size(600D, 800D)
            .Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Sensitive review text")))))).ToBytes();
        byte[] annotated = PdfDocument.Open(source).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Text",
            Rectangle = [40D, 50D, 60D, 70D],
            Contents = "Original",
            Title = "Reviewer"
        }).Bytes;
        await File.WriteAllBytesAsync(path, annotated);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);
            int objectNumber = Assert.Single(PdfDocument.Open(annotated).Inspect().GetAnnotationsBySubtype("Text")).ObjectNumber!.Value;
            viewModel.Pages[0].SelectAnnotation(new PdfEditorSelection(1, objectNumber, "Text"));
            Assert.True(viewModel.HasSelectedAnnotation);
            viewModel.SelectedEditorToolChoice = viewModel.EditorTools.Single(choice => choice.Tool == PdfEditorTool.Redact);
            viewModel.Pages[0].CompleteEditorGesture(new PdfEditorGesture(1, 36D, 48D, 240D, 92D, Array.Empty<PdfEditorVisualPoint>()));
            await WaitUntilAsync(() => viewModel.HasPendingRedaction);

            viewModel.SetOrganizerSelection([viewModel.OrganizerPages[0]]);
            await viewModel.DuplicateSelectedCommand.ExecuteAsync(null);

            Assert.False(viewModel.HasPendingRedaction);
            Assert.False(viewModel.HasSelectedAnnotation);
            Assert.All(viewModel.Pages, page => Assert.Null(page.PendingRedactionArea));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task AnnotationUpdatePreservesEditedContentsAndAuthor() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-annotation-update-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "editable.pdf");
        byte[] source = PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D))).ToBytes();
        byte[] annotated = PdfDocument.Open(source).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Text",
            Rectangle = [40D, 50D, 60D, 70D],
            Contents = "Original",
            Title = "Reviewer"
        }).Bytes;
        await File.WriteAllBytesAsync(path, annotated);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);
            int objectNumber = Assert.Single(PdfDocument.Open(annotated).Inspect().GetAnnotationsBySubtype("Text")).ObjectNumber!.Value;
            viewModel.Pages[0].SelectAnnotation(new PdfEditorSelection(1, objectNumber, "Text"));
            viewModel.SelectedAnnotationContents = "Edited contents";
            viewModel.SelectedAnnotationAuthor = "Edited author";

            await viewModel.UpdateSelectedAnnotationCommand.ExecuteAsync(null);
            await viewModel.SaveCommand.ExecuteAsync(null);

            PdfAnnotation updated = Assert.Single(PdfDocument.Open(path).Inspect().GetAnnotationsBySubtype("Text"));
            Assert.Equal("Edited contents", updated.Contents);
            Assert.Equal("Edited author", updated.Title);
            Assert.False(viewModel.HasSelectedAnnotation);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task AddImagePromptsForEachPlacementAndDoesNotReuseRetainedBytes() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-image-picker-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "editable.pdf");
        string imagePath = Path.Combine(root, "pixel.png");
        PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D))).Save(path);
        await File.WriteAllBytesAsync(imagePath, TinyPng);
        int picks = 0;

        try {
            using var viewModel = new MainWindowViewModel(
                _ => Task.FromResult<string?>(null),
                pickImage: _ => {
                    picks++;
                    return Task.FromResult<string?>(imagePath);
                });
            await viewModel.OpenDocumentAsync(path);
            viewModel.SelectedEditorToolChoice = viewModel.EditorTools.Single(choice => choice.Tool == PdfEditorTool.AddImage);
            var gesture = new PdfEditorGesture(1, 40D, 50D, 80D, 90D, Array.Empty<PdfEditorVisualPoint>());
            PdfPageViewModel firstPage = viewModel.Pages[0];
            firstPage.CompleteEditorGesture(gesture);
            await WaitUntilAsync(() => !ReferenceEquals(firstPage, viewModel.Pages[0]));
            PdfPageViewModel secondPage = viewModel.Pages[0];
            secondPage.CompleteEditorGesture(gesture with { Left = 100D, Right = 140D });
            await WaitUntilAsync(() => !ReferenceEquals(secondPage, viewModel.Pages[0]));

            Assert.Equal(2, picks);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task FailedEditorGestureKeepsFailureStatus() {
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        await viewModel.OpenDocumentAsync(GetFixturePath());
        viewModel.SelectedEditorToolChoice = viewModel.EditorTools.Single(choice => choice.Tool == PdfEditorTool.AddText);

        viewModel.Pages[0].CompleteEditorGesture(new PdfEditorGesture(1, 40D, 50D, 180D, 100D, Array.Empty<PdfEditorVisualPoint>()));
        await WaitUntilAsync(() => !viewModel.IsWorkspaceBusy && viewModel.HasError);

        Assert.Equal("Operation failed", viewModel.OperationStatus);
        Assert.DoesNotContain("Edit added", viewModel.OperationStatus, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AppendOnlyFormFillDoesNotEnableFillAndFlatten() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-form-capability-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "certified-form.pdf");
        byte[] unsigned = PdfDocument.Create(compose => compose.Page(page => page.Content(content => content.Item(item =>
            item.TextField("Customer.Name", value: "Before"))))).ToBytes();
        PdfExternalSignaturePreparation preparation = PdfDocument.Open(unsigned).Security.PrepareExternalSignature(new PdfExternalSignatureOptions {
            Profile = PdfSignatureProfile.Certification,
            CertificationPermission = PdfCertificationPermissionLevel.FormFillingAndSignatures,
            FieldName = "Certification",
            ReservedSignatureContentsBytes = 512
        });
        await File.WriteAllBytesAsync(path, preparation.Complete([0x30, 0x01, 0x00]).ToBytes());

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);

            Assert.True(viewModel.CanFillForms);
            Assert.False(viewModel.CanFlattenForms);
            Assert.False(viewModel.CanFillAndFlattenForms);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task OpenCommandLoadsPathReturnedByPicker() {
        string fixture = GetFixturePath();
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(fixture));

        await viewModel.OpenCommand.ExecuteAsync(null);

        Assert.True(viewModel.HasDocument);
        Assert.False(viewModel.IsEmpty);
        Assert.True(viewModel.IsPdfWorkspaceMode);
        Assert.False(viewModel.IsHomeMode);
        Assert.Equal(System.IO.Path.GetFileName(fixture), viewModel.DocumentName);
        Assert.NotEmpty(viewModel.Pages);
        Assert.Same(viewModel.Pages[0], viewModel.SelectedPage);
    }

    [Fact]
    public async Task FailedOpenProducesDismissibleErrorState() {
        string missing = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            $"officeimo-studio-missing-{Guid.NewGuid():N}.pdf");
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(missing));

        await viewModel.OpenCommand.ExecuteAsync(null);

        Assert.False(viewModel.HasDocument);
        Assert.True(viewModel.IsEmpty);
        Assert.True(viewModel.IsHomeMode);
        Assert.True(viewModel.HasError);
        Assert.Contains("no longer exists", viewModel.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        viewModel.DismissErrorCommand.Execute(null);
        Assert.False(viewModel.HasError);
    }

    [Fact]
    public void FitPageRecomputesWhenSelectedPageDimensionsChange() {
        using var coordinator = new PageRenderCoordinator((page, scale, _) =>
            Task.FromResult(new PdfRenderedPage(
                page,
                scale,
                [1],
                1,
                1,
                TimeSpan.Zero,
                Array.Empty<string>())));
        using var sceneCoordinator = new PageSceneCoordinator((page, _) =>
            Task.FromResult(TestPdfPageScenes.Create(page)));
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        var squarePage = new PdfPageViewModel(1, 400, 400, 0, 1D, sceneCoordinator, coordinator);
        var tallPage = new PdfPageViewModel(2, 400, 1000, 0, 1D, sceneCoordinator, coordinator);
        viewModel.Pages.Add(squarePage);
        viewModel.Pages.Add(tallPage);
        viewModel.SetViewportSize(1000, 700);
        viewModel.SelectedPage = squarePage;

        viewModel.FitPageCommand.Execute(null);
        double squareZoom = viewModel.Zoom;
        viewModel.SelectedPage = tallPage;

        Assert.True(viewModel.Zoom < squareZoom);
        Assert.Equal(0.63D, viewModel.Zoom);
    }

    [Fact]
    public async Task OrganizerMutationRefreshesReaderAndSupportsUndo() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-shell-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "editable.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(500, 700))).Save(path);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);
            viewModel.SetOrganizerSelection([viewModel.OrganizerPages[0]]);

            await viewModel.DuplicateSelectedCommand.ExecuteAsync(null);

            Assert.Equal(2, viewModel.Pages.Count);
            Assert.True(viewModel.IsDirty);
            Assert.True(viewModel.CanUndo);

            await viewModel.UndoCommand.ExecuteAsync(null);
            Assert.Single(viewModel.Pages);
            Assert.False(viewModel.IsDirty);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task MultiplePickerImportsAsOneCommandAndSelectsImportedPages() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-shell-multi-import-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string target = Path.Combine(root, "target.pdf");
        string first = Path.Combine(root, "first.pdf");
        string second = Path.Combine(root, "second.pdf");
        CreateTextSource(target, "Target");
        CreateTextSource(first, "First");
        CreateTextSource(second, "Second A", "Second B");

        try {
            using var viewModel = new MainWindowViewModel(
                _ => Task.FromResult<string?>(null),
                pickImportPdfs: _ => Task.FromResult<IReadOnlyList<string>>([first, second]));
            await viewModel.OpenDocumentAsync(target);

            await viewModel.ImportPagesCommand.ExecuteAsync(null);

            Assert.Equal(4, viewModel.Pages.Count);
            Assert.Equal([2, 3, 4], viewModel.OrganizerPages.Where(page => page.IsSelected).Select(page => page.PageNumber).ToArray());
            Assert.Equal("Imported 3 pages from 2 PDFs", viewModel.OperationStatus);

            await viewModel.UndoCommand.ExecuteAsync(null);
            Assert.Single(viewModel.Pages);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task PageDeletionRequiresConfirmationAndRemainsUndoable() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-delete-confirm-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "two-pages.pdf");
        CreateTextSource(path, "One", "Two");
        bool allowDeletion = false;
        int confirmationCount = 0;

        try {
            using var viewModel = new MainWindowViewModel(
                _ => Task.FromResult<string?>(null),
                confirmPageDeletion: count => {
                    confirmationCount += count;
                    return Task.FromResult(allowDeletion);
                });
            await viewModel.OpenDocumentAsync(path);
            viewModel.SetOrganizerSelection([viewModel.OrganizerPages[0]]);

            await viewModel.DeleteSelectedCommand.ExecuteAsync(null);

            Assert.Equal(2, viewModel.Pages.Count);
            Assert.False(viewModel.IsDirty);
            Assert.Equal("Delete cancelled", viewModel.OperationStatus);
            Assert.Equal(1, confirmationCount);

            allowDeletion = true;
            await viewModel.DeleteSelectedCommand.ExecuteAsync(null);

            Assert.Single(viewModel.Pages);
            Assert.True(viewModel.IsDirty);
            await viewModel.UndoCommand.ExecuteAsync(null);
            Assert.Equal(2, viewModel.Pages.Count);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task OrganizerSelectionTracksMovedPagesAndSurvivesRotation() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-organizer-selection-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "three-pages.pdf");
        CreateTextSource(path, "One", "Two", "Three");

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);
            viewModel.SetOrganizerSelection([viewModel.OrganizerPages[1], viewModel.OrganizerPages[2]]);

            await viewModel.MoveSelectedUpCommand.ExecuteAsync(null);

            Assert.Equal([1, 2], viewModel.OrganizerPages.Where(page => page.IsSelected).Select(page => page.PageNumber).ToArray());
            await viewModel.RotateRightCommand.ExecuteAsync(null);
            Assert.Equal([1, 2], viewModel.OrganizerPages.Where(page => page.IsSelected).Select(page => page.PageNumber).ToArray());
            Assert.All(viewModel.OrganizerPages.Take(2), page => Assert.Equal(90, page.RotationDegrees));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task SelectAllAndClearRemainStableAcrossVirtualizedSelectionEvents() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-organizer-virtualization-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "five-pages.pdf");
        CreateTextSource(path, "One", "Two", "Three", "Four", "Five");

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);

            viewModel.SelectAllPagesCommand.Execute(null);
            viewModel.UpdateOrganizerSelection(viewModel.OrganizerPages.Take(2), []);

            Assert.Equal("5 of 5 selected", viewModel.OrganizerSelectionLabel);
            Assert.All(viewModel.OrganizerPages, page => Assert.True(page.IsSelected));

            viewModel.ClearPageSelectionCommand.Execute(null);
            viewModel.UpdateOrganizerSelection([], viewModel.OrganizerPages.Take(2));

            Assert.Equal("Select pages", viewModel.OrganizerSelectionLabel);
            Assert.All(viewModel.OrganizerPages, page => Assert.False(page.IsSelected));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task OrganizerActivationNavigatesWithoutChangingBulkSelection() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-organizer-navigation-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "three-pages.pdf");
        CreateTextSource(path, "One", "Two", "Three");

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);
            viewModel.SetOrganizerSelection([viewModel.OrganizerPages[0], viewModel.OrganizerPages[2]]);
            viewModel.NavigateToOrganizerPage(2);

            Assert.Equal(2, viewModel.SelectedPage?.PageNumber);
            Assert.Equal([1, 3], viewModel.OrganizerPages.Where(page => page.IsSelected).Select(page => page.PageNumber).ToArray());

            viewModel.NavigateToOrganizerPage(1);

            Assert.Equal(1, viewModel.SelectedPage?.PageNumber);
            Assert.Equal([1, 3], viewModel.OrganizerPages.Where(page => page.IsSelected).Select(page => page.PageNumber).ToArray());
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task SplitCommandUsesConfiguredPagesPerDocument() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-shell-split-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "three-pages.pdf");
        string output = Path.Combine(root, "output");
        CreateTextSource(path, "One", "Two", "Three");

        try {
            using var viewModel = new MainWindowViewModel(
                _ => Task.FromResult<string?>(null),
                pickOutputFolder: _ => Task.FromResult<string?>(output));
            await viewModel.OpenDocumentAsync(path);
            viewModel.SplitPagesPerDocument = 2;

            await viewModel.SplitCommand.ExecuteAsync(null);

            Assert.Equal(2, Directory.GetFiles(output, "*.pdf").Length);
            Assert.Equal("Created 2 split PDFs", viewModel.OperationStatus);
            Assert.False(viewModel.IsDirty);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task BlockedOrganizerMutationDoesNotReportSuccessOrRefreshTheDocument() {
        string fixture = GetFixturePath();
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        await viewModel.OpenDocumentAsync(fixture);
        PdfPageViewModel originalPage = viewModel.Pages[0];
        viewModel.SetOrganizerSelection([viewModel.OrganizerPages[0]]);

        await viewModel.DuplicateSelectedCommand.ExecuteAsync(null);

        Assert.False(viewModel.CanMutatePages);
        Assert.Same(originalPage, viewModel.Pages[0]);
        Assert.Single(viewModel.Pages);
        Assert.Equal("Operation failed", viewModel.OperationStatus);
        Assert.True(viewModel.HasError);
    }

    [Fact]
    public async Task DiscardedDirtyDocumentDeletesRecoveryBeforeOpeningReplacement() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-transition-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string first = Path.Combine(root, "first.pdf");
        string second = Path.Combine(root, "second.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(500, 700))).Save(first);
        PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).Save(second);

        try {
            using (var viewModel = new MainWindowViewModel(
                _ => Task.FromResult<string?>(null),
                confirmUnsavedChanges: () => Task.FromResult(UnsavedChangesDecision.Discard))) {
                await viewModel.OpenDocumentAsync(first);
                viewModel.SetOrganizerSelection([viewModel.OrganizerPages[0]]);
                await viewModel.DuplicateSelectedCommand.ExecuteAsync(null);
                Assert.True(viewModel.IsDirty);

                await viewModel.OpenDocumentAsync(second);

                Assert.Equal("second.pdf", viewModel.DocumentName);
                Assert.Single(viewModel.Pages);
                Assert.False(viewModel.IsDirty);
            }

            using PdfWorkspace reopened = await PdfWorkspace.OpenAsync(first, CancellationToken.None);
            Assert.False(reopened.HasRecovery);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task DroppingOntoAnotherSelectedPageIsANoOp() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-drop-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "four-pages.pdf");
        PdfDocument.Create(compose => {
            for (int page = 0; page < 4; page++) compose.Page(item => item.Size(500 + page, 700));
        }).Save(path);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);
            PdfPageViewModel[] original = viewModel.Pages.ToArray();
            viewModel.SetOrganizerSelection([viewModel.OrganizerPages[1], viewModel.OrganizerPages[2]]);

            await viewModel.ReorderByDropAsync(2, 3);

            Assert.False(viewModel.IsDirty);
            Assert.Equal(original, viewModel.Pages);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task StandardNamedPageActionsNavigateInsideTheDocument() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-links-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "links.pdf");
        PdfDocument.Create(compose => {
            compose.Page(page => page.Size(500, 700));
            compose.Page(page => page.Size(500, 700));
        }).Save(path);

        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);

            await viewModel.ActivatePageLinkAsync("LastPage");
            Assert.Equal(2, viewModel.SelectedPage?.PageNumber);
            await viewModel.ActivatePageLinkAsync("PrevPage");
            Assert.Equal(1, viewModel.SelectedPage?.PageNumber);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task CancellingDirtyDocumentTransitionKeepsTheCurrentWorkspace() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-transition-cancel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string first = Path.Combine(root, "first.pdf");
        string second = Path.Combine(root, "second.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(500, 700))).Save(first);
        PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).Save(second);
        UnsavedChangesDecision decision = UnsavedChangesDecision.Cancel;

        try {
            using var viewModel = new MainWindowViewModel(
                _ => Task.FromResult<string?>(null),
                confirmUnsavedChanges: () => Task.FromResult(decision));
            await viewModel.OpenDocumentAsync(first);
            viewModel.SetOrganizerSelection([viewModel.OrganizerPages[0]]);
            await viewModel.DuplicateSelectedCommand.ExecuteAsync(null);

            await viewModel.OpenDocumentAsync(second);

            Assert.Equal("first.pdf *", viewModel.DocumentName);
            Assert.Equal(2, viewModel.Pages.Count);
            Assert.True(viewModel.IsDirty);

            decision = UnsavedChangesDecision.Discard;
            Assert.True(await viewModel.RequestCloseDocumentAsync());
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ClosingDocumentClearsDocumentScopedSearchAndStatus() {
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        await viewModel.OpenDocumentAsync(GetFixturePath());
        viewModel.SearchQuery = "test";
        await viewModel.SearchCommand.ExecuteAsync(null);
        Assert.NotEmpty(viewModel.SearchResults);
        Assert.True(viewModel.HasOperationStatus);

        Assert.True(await viewModel.RequestCloseDocumentAsync());

        Assert.False(viewModel.HasDocument);
        Assert.Empty(viewModel.SearchQuery);
        Assert.Empty(viewModel.SearchResults);
        Assert.Null(viewModel.SelectedSearchResult);
        Assert.False(viewModel.HasOperationStatus);
        Assert.Equal(0D, viewModel.OperationProgressFraction);
    }

    private static string GetFixturePath() =>
        System.IO.Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf");

    private static void CreateTextSource(string path, params string[] pageTexts) =>
        PdfDocument.Create(compose => {
            foreach (string text in pageTexts) {
                compose.Page(page => page.Content(content => content.Item(item =>
                    item.Paragraph(paragraph => paragraph.Text(text)))));
            }
        }).Save(path);

    private static async Task WaitUntilAsync(Func<bool> condition) {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        while (!condition()) await Task.Delay(10, timeout.Token);
    }

    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
}

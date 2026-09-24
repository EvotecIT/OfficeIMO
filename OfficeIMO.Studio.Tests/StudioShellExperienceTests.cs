using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioShellExperienceTests {
    [Fact]
    public async Task OperationToastOffersUndoOnlyForItsOwnEdit() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            string source = Path.Combine(services.Paths.Root, "toast.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(400, 500))).Save(source);
            var window = new MainWindow(services) { Width = 960, Height = 620 };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(source);
                var document = window.ViewModel;
                document.SetOrganizerSelection([document.OrganizerPages[0]]);
                await document.DuplicateSelectedCommand.ExecuteAsync(null);
                window.UpdateLayout();
                var undo = window.FindControl<Button>("ToastUndoButton")!;
                Assert.True(undo.IsVisible);
                CaptureToast(window, "edit");

                document.OperationStatus = "Exported to copy.pdf";
                window.UpdateLayout();
                Assert.True(document.CanUndo);
                Assert.False(undo.IsVisible);
                CaptureToast(window, "export");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void CaptureToast(MainWindow window, string state) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        frame.Save(Path.Combine(output, $"operation-toast-{state}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }

    [Fact]
    public async Task CommandSearchListsRunnableAndRecentCommandsBeforeUnavailableOnes() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            var palette = new StudioCommandPaletteModel(model.Commands);
            IReadOnlyList<StudioCommandItem> results = palette.Results;
            int lastAvailable = results.Select((item, index) => (item, index)).Last(entry => entry.item.IsAvailable).index;
            int firstUnavailable = results.Select((item, index) => (item, index)).First(entry => !entry.item.IsAvailable).index;
            Assert.True(lastAvailable < firstUnavailable);
            Assert.Same(results[0], palette.SelectedCommand);

            model.Commands.MarkUsed("Settings");
            palette.Refresh();
            Assert.Equal("Settings", palette.Results[0].Id);
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData("Rotate selected pages", "rot", new[] { 0, 3 })]
    [InlineData("Save a copy", "save copy", new[] { 0, 4, 7, 4 })]
    [InlineData("Protect and sign", "zzz", new int[0])]
    public void CommandSearchHighlightsEachMatchedTermOnce(string text, string query, int[] expected) {
        int[] flattened = StudioTextHighlight.FindMatches(text, query).SelectMany(match => new[] { match.Start, match.Length }).ToArray();
        Assert.Equal(expected, flattened);
    }

    [Theory]
    [InlineData("render.resource.font-substitution", "Font substitution")]
    [InlineData("RouteContract", "Route contract")]
    [InlineData("FixedLayout", "Fixed layout")]
    [InlineData("PDFExport", "PDF export")]
    public void EngineCodesReadAsPlainWords(string code, string expected) =>
        Assert.Equal(expected, StudioMessages.Humanize(code));

    [Fact]
    public void ArgumentMessagesDropTheParameterName() =>
        Assert.Equal("The selected printer is not installed.",
            StudioMessages.Describe(new ArgumentException("The selected printer is not installed.", "printerName")));

    [Fact]
    public void CertifyingSelectsTheCertificationProfileAndItsAllowedChanges() {
        var certificate = new PdfSigningCertificateViewModel("AB", "Signer", DateTime.Today.AddYears(1), "Issuer");
        var approval = new PdfSigningSettings(certificate, "Signature1", "Approved", "", false, 1, 0, 0, 100, 40).CreateOptions();
        var certification = new PdfSigningSettings(certificate, "Signature1", "Approved", "", false, 1, 0, 0, 100, 40,
            PdfCertificationPermissionLevel.FormFillingAndSignatures).CreateOptions();
        Assert.Equal(PdfSignatureProfile.Approval, approval.Profile);
        Assert.Equal(PdfSignatureProfile.Certification, certification.Profile);
        Assert.Equal(PdfCertificationPermissionLevel.FormFillingAndSignatures, certification.CertificationPermission);
    }

    [Fact]
    public async Task ReopenClosedTabRestoresTheMostRecentlyClosedDocument() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-reopen-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "closed.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).Save(path);
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var window = new MainWindow();
                try {
                    window.Show();
                    await window.TabHost.OpenDocumentAsync(path);
                    var tab = Assert.Single(window.TabHost.Tabs);
                    await window.TabHost.CloseTabAsync(tab);
                    Assert.Empty(window.TabHost.Tabs);
                    Assert.True(window.TabHost.CanReopenClosedTab);

                    await window.TabHost.ReopenClosedTabAsync();
                    var reopened = Assert.Single(window.TabHost.Tabs);
                    Assert.Equal("closed.pdf", reopened.DisplayTitle);
                    Assert.False(window.TabHost.CanReopenClosedTab);
                } finally {
                    window.Close();
                }
                return true;
            }, CancellationToken.None);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task SelectionMarkupSwitchesToAnnotateAndCreatesAnUndoableHighlight() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-selection-markup-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "markup.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D)
            .Content(content => content.Text("Highlight this sentence from the reader.")))).Save(path);
        try {
            using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await viewModel.OpenDocumentAsync(path);
            Assert.Equal(StudioDocumentMode.View, viewModel.DocumentMode);
            Assert.False(viewModel.CanUndo);

            viewModel.Pages[0].RequestMarkup(PdfEditorTool.Highlight, new PdfEditorGesture(1, 36D, 40D, 320D, 70D,
                [new PdfEditorVisualPoint(36D, 40D), new PdfEditorVisualPoint(320D, 70D)]));
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            while (!viewModel.CanUndo) await Task.Delay(10, timeout.Token);

            Assert.Equal(StudioDocumentMode.Annotate, viewModel.DocumentMode);
            Assert.Equal(PdfEditorTool.Highlight, viewModel.ActiveEditorTool);
            Assert.False(viewModel.HasError, viewModel.ErrorMessage);
            Assert.True(viewModel.IsDirty);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }
}

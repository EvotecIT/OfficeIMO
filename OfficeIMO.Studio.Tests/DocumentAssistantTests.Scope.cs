using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed partial class DocumentAssistantTests {
    [Theory]
    [InlineData(1)]
    [InlineData(501)]
    public async Task CurrentPagePreparationReadsOnlyTheSelectedPageAndKeepsOriginalNumber(int selectedPage) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            Directory.CreateDirectory(services.Paths.Root);
            string path = Path.Combine(services.Paths.Root, "large-scope.pdf");
            PdfDocument.Create(document => {
                for (int page = 1; page <= 501; page++) {
                    int number = page;
                    document.Page(item => item.Content(content => content.Text($"Source page {number}.")));
                }
            }).Save(path);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await model.OpenDocumentAsync(path);
            Assert.True(model.HasDocument, model.ErrorMessage);
            model.SelectedPage = model.Pages[selectedPage - 1];
            model.Assistant.CurrentPageOnly = true;
            await model.Assistant.PrepareEvidenceCommand.ExecutionTask!;
            var prepared = model.Assistant.PreparedDocument;
            Assert.NotNull(prepared);
            Assert.Equal(selectedPage, Assert.Single(prepared.Pages));
            Assert.All(prepared.Evidence, item => Assert.Equal(selectedPage, item.Page));
            Assert.Contains(prepared.Evidence, item => item.Text.Contains($"Source page {selectedPage}."));
            Assert.Equal(File.ReadAllBytes(path).Length, prepared.SourceByteLength);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ChangingTabsClosesTheInactiveAssistantAndReopeningPreparesEvidence() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            Directory.CreateDirectory(services.Paths.Root);
            string first = Path.Combine(services.Paths.Root, "first.pdf");
            string second = Path.Combine(services.Paths.Root, "second.pdf");
            PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("First document evidence.")))).Save(first);
            PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("Second document evidence.")))).Save(second);
            var window = new MainWindow(services) { Width = 960, Height = 620 };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(first);
                var firstTab = window.TabHost.SelectedTab!;
                window.ViewModel.ToggleAssistantCommand.Execute(null);
                await window.ViewModel.Assistant.PrepareEvidenceCommand.ExecutionTask!;
                Assert.NotNull(firstTab.Document.Assistant.PreparedDocument);
                await window.TabHost.OpenDocumentAsync(second);
                Assert.False(firstTab.Document.IsAssistantVisible);
                Assert.Null(firstTab.Document.Assistant.PreparedDocument);
                window.TabHost.SelectedTab = firstTab;
                Assert.False(window.FindControl<SplitView>("AssistantHost")!.IsPaneOpen);
                window.ViewModel.ToggleAssistantCommand.Execute(null);
                await window.ViewModel.Assistant.PrepareEvidenceCommand.ExecutionTask!;
                Assert.NotNull(firstTab.Document.Assistant.PreparedDocument);
                Assert.True(window.FindControl<SplitView>("AssistantHost")!.IsPaneOpen);
                Assert.Contains(firstTab.Document.Assistant.PreparedDocument!.Evidence, item => item.Text.Contains("First document evidence."));
                window.UpdateLayout(); Dispatcher.UIThread.RunJobs(); AvaloniaHeadlessPlatform.ForceRenderTimerTick();
                using var frame = window.CaptureRenderedFrame();
                Assert.NotNull(frame);
                string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (output is not null) {
                    Directory.CreateDirectory(output);
                    frame.Save(Path.Combine(output, "assistant-tab-reopened-960x620.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
}
